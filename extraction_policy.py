"""Shared conservative handling for model responses and draft training labels."""

POLICY_VERSION = "explicit-evidence-v1"


def mask_uncertain_fields(fields, uncertain_fields):
    """Keep uncertain readings outside extraction targets, preserving review candidates."""
    values = dict(fields)
    uncertain = {key for key in uncertain_fields if isinstance(key, str) and key in values}
    if uncertain & {"surname", "given_name"}:
        uncertain.update({"surname", "given_name"} & values.keys())
    candidates = {key: values[key] for key in sorted(uncertain) if values[key]}
    for key in uncertain:
        values[key] = ""
    return values, sorted(uncertain), candidates
