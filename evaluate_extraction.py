"""Compare prompts locally against explicitly human-verified image/JSON references."""
import argparse
import json
import hashlib
import time
from pathlib import Path

from PIL import Image
import extract_decisions as app


def score_predictions(references, predictions):
    """Score reviewed fields only; report hallucinated fills separately from omissions."""
    predictions_by_id = {row["record_id"]: row["output"] for row in predictions}
    scores = {key: {"compared": 0, "correct": 0, "nonempty": 0,
                    "nonempty_correct": 0, "unsupported_fills": 0, "omissions": 0}
              for key in app.FIELDS}
    verified = [row for row in references if row.get("human_verified") is True]
    if not verified:
        raise ValueError("No human-verified references. Review images and labels first; no model calls made.")
    for row in verified:
        expected = row["output"]
        actual = predictions_by_id[row["record_id"]]
        for key in app.FIELDS:
            if key in expected.get("uncertain_fields", []):
                continue
            want, got = expected[key], actual.get(key, "")
            count = scores[key]
            count["compared"] += 1
            # Exact text comparison after whitespace cleanup. Dates retain their printed form.
            equal = app.clean_value(want) == app.clean_value(got)
            count["correct"] += equal
            count["nonempty"] += bool(want)
            count["nonempty_correct"] += bool(want) and equal
            count["unsupported_fills"] += not want and bool(got)
            count["omissions"] += bool(want) and not got
    totals = {key: sum(count[key] for count in scores.values()) for key in next(iter(scores.values()))}
    return {"human_verified_images": len(verified), "counts": totals, "fields": scores,
            "exact_field_accuracy": totals["correct"] / totals["compared"] if totals["compared"] else None,
            "nonempty_field_accuracy": totals["nonempty_correct"] / totals["nonempty"] if totals["nonempty"] else None}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--references", type=Path, required=True, help="JSONL with record_id, image, output, human_verified")
    parser.add_argument("--image-root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--model", default=app.DEFAULT_MODEL)
    parser.add_argument("--baseline-prompt", type=Path)
    parser.add_argument("--baseline-crop-prompt", type=Path)
    parser.add_argument("--max-side", type=int, default=app.MAX_SIDE)
    parser.add_argument("--num-ctx", type=int, default=app.NUM_CTX)
    parser.add_argument("--request-timeout", type=int, default=3600,
                        help="Seconds per image; CPU vision processing can take longer than the app's default")
    args = parser.parse_args()
    if args.request_timeout < 1:
        parser.error("--request-timeout must be positive")
    app.TIMEOUT_SECONDS = args.request_timeout
    references = [json.loads(line) for line in args.references.read_text(encoding="utf-8-sig").splitlines() if line.strip()]
    references = [row for row in references if row.get("human_verified") is True]
    if not references:
        parser.error("No human-verified references. Review images and labels first; no model calls made.")
    if len({row["record_id"] for row in references}) != len(references):
        parser.error("Duplicate reference record IDs")
    for row in references:
        if row.get("exclude_training") or row.get("duplicate_of_record_id"):
            parser.error("Use independent references without excluded documents or duplicate photographs")
        if set(row["output"]) != set(app.SCHEMA["required"]):
            parser.error("Reference output does not match the response schema")
        if not (args.image_root / row["image"]).is_file():
            parser.error(f"Missing reference image: {row['image']}")
    app.check_ollama(args.model)
    candidate_system, candidate_crop = app.SYSTEM_PROMPT, app.CROP_PROMPT
    variants = [("candidate", candidate_system, candidate_crop)]
    if args.baseline_prompt:
        baseline_crop = args.baseline_crop_prompt.read_text(encoding="utf-8") if args.baseline_crop_prompt else candidate_crop
        variants.insert(0, ("baseline", args.baseline_prompt.read_text(encoding="utf-8"), baseline_crop))
    reports = {}
    args.output.mkdir(parents=True, exist_ok=True)
    manifest = {"model": args.model, "max_side": args.max_side, "num_ctx": args.num_ctx,
                "references": hashlib.sha256(json.dumps(references, ensure_ascii=False, sort_keys=True).encode()).hexdigest(),
                "prompts": {name: hashlib.sha256(system.encode()).hexdigest() for name, system, _ in variants}}
    manifest_path = args.output / "run_manifest.json"
    if manifest_path.exists() and json.loads(manifest_path.read_text(encoding="utf-8")) != manifest:
        parser.error("Output folder belongs to a different evaluation configuration; choose another --output folder")
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    (args.output / "references.jsonl").write_text("".join(json.dumps(row, ensure_ascii=False) + "\n" for row in references), encoding="utf-8")
    try:
        for name, system, crop in variants:
            app.SYSTEM_PROMPT, app.CROP_PROMPT = system, crop
            path = args.output / f"{name}.predictions.jsonl"
            predictions = [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()] if path.exists() else []
            completed = {row["record_id"] for row in predictions}
            (args.output / f"{name}.prompt.txt").write_text(system, encoding="utf-8")
            for position, row in enumerate(references, 1):
                if row["record_id"] in completed:
                    continue
                print(f"{name} {position}/{len(references)}: {row['image']}", flush=True)
                started = time.monotonic()
                with Image.open(args.image_root / row["image"]) as original:
                    # Reference images are already upright; no retries or cache reuse.
                    result, _, _, _ = app.model_page(original.convert("RGB"), args.model,
                                                    args.max_side, args.num_ctx, 0, False)
                predictions.append({"record_id": row["record_id"], "output": result["fields"],
                                    "uncertain_fields": result["uncertain_fields"],
                                    "elapsed_seconds": round(time.monotonic() - started, 2)})
                app.atomic_json(args.output / f"{name}.progress.json", {"completed": len(predictions), "total": len(references)})
                path.write_text("".join(json.dumps(item, ensure_ascii=False) + "\n" for item in predictions), encoding="utf-8")
                print(f"saved {name} {position}/{len(references)} in {predictions[-1]['elapsed_seconds']}s", flush=True)
            (args.output / f"{name}.predictions.jsonl").write_text(
                "".join(json.dumps(row, ensure_ascii=False) + "\n" for row in predictions), encoding="utf-8")
            reports[name] = score_predictions(references, predictions)
    finally:
        app.SYSTEM_PROMPT, app.CROP_PROMPT = candidate_system, candidate_crop
    report = {"model": args.model, "max_side": args.max_side, "num_ctx": args.num_ctx,
              "detail": False, "comparison": "same pipeline and images; system prompt differs; crop pass disabled",
              "results": reports}
    (args.output / "comparison.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({name: {k: v for k, v in score.items() if k != "fields"} for name, score in reports.items()}))


if __name__ == "__main__":
    main()
