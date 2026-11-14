#!/usr/bin/env python3

# -*- coding: utf-8 -*-

"""Offline Arabic administrative-decision extractor for scanned PDFs and images.



Requires Ollama with a local vision model, PyMuPDF, Pillow and openpyxl.

No text is read from any PDF text layer: every page is rendered as an image.



Quick start:

    python extract_decisions.py --limit 2

    python extract_decisions.py

    python extract_decisions.py --detail

    python extract_decisions.py --model qwen3.5:9b



Only official document images are transmitted to http://localhost:11434 (Ollama).

They are NEVER uploaded to a remote OCR service by this application.

"""



from __future__ import annotations



import argparse

import base64

import csv

import datetime as dt

import hashlib

import io

import json

import re

import sys

import time
import threading

import unicodedata

import urllib.error

import urllib.request

from collections import defaultdict

from dataclasses import dataclass, field

from pathlib import Path

from typing import Any



from PIL import Image, ImageOps



try:

    import pymupdf  # Current PyMuPDF import name

except ImportError:

    import fitz as pymupdf  # Older PyMuPDF versions



from openpyxl import Workbook

from openpyxl.styles import Alignment, Font, PatternFill

from openpyxl.utils import get_column_letter





# --------------------- EDITABLE SETTINGS ---------------------

DEFAULT_MODEL = "qwen3.5:9b"  # smaller, CPU-friendly quantized model

OLLAMA_URL = "http://127.0.0.1:11434"  # explicitly local, not a cloud service

PDF_DPI = 200

MAX_SIDE = 1700  # pixels sent to model (not the PDF render resolution)

NUM_CTX = 4096

NUM_PREDICT = 1600

TIMEOUT_SECONDS = 900

JPEG_QUALITY = 91

CACHE_VERSION = "arabic-decisions-v3-2026-11-03"

IMG_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"}



# Each model key has the final Arabic Excel column on the right.

# Add new columns here AND a specific instruction to the prompt to extend the project.

FIELDS: dict[str, str] = {

    "decision_number": "رقم المقرر",

    "decision_date": "تاريخ المقرر",

    "decision_effective_date": "تاريخ سريان المقرر",

    "surname": "اللقب",

    "given_name": "الإسم",

    "full_name_as_written": "الاسم الكامل كما ورد",

    "gender": "الجنس",

    "category": "الصنف",

    "decision_type": "نوع المقرر",

    "decision_subject": "موضوع المقرر",

    "current_rank": "الرتبة الحالية",

    "branch": "الشعبة",

    "appointment_place": "مكان التعيين",

    "old_rank": "الرتبة قبل الترقية",

    "new_rank": "الرتبة بعد الترقية",

    "current_rank_appointment_date": "تاريخ التعيين في الرتبة الحالية",

    "grade": "الدرجة",

    "grade_effective_date": "تاريخ سريان الدرجة",

    "senior_post": "المنصب العالي",

    "senior_post_appointment_date": "تاريخ التعيين في المنصب العالي",

    "higher_function": "الوظيفة العليا",

    "higher_function_appointment_date": "تاريخ التعيين في الوظيفة العليا",

    "administrative_movement": "الحركة الإدارية",

    "movement_date": "تاريخ الحركة",

    "visa_number": "رقم تأشيرة",

    "visa_date": "تاريخ التأشيرة",

    "employment_status": "الصفة المهنية (مرسم/متعاقد/متربص)",

    "legal_status": "الوضعية القانونية",

    "service_status": "وضعية الخدمة",

    "index_number": "الرقم الاستدلالي",

    "organization": "الهيئة أو المصلحة",

    "notes": "ملاحظات",

}



DATE_FIELDS = {

    "decision_date", "decision_effective_date", "current_rank_appointment_date",

    "grade_effective_date", "senior_post_appointment_date", "higher_function_appointment_date",

    "movement_date", "visa_date",

}

IDENTITY_FIELDS = ("decision_number", "decision_date", "surname", "given_name", "decision_type")

KEY_FIELDS = ("decision_number", "decision_date", "surname", "given_name", "decision_type")

BLANK_MARKERS = {"null", "none", "n/a", "na", "unknown", "-", "--", "غير معروف", "غير متوفر", "غير موجود"}

PAGE_ROLES = {"start", "continuation", "unknown"}



# Return an explicit page role so adjacent pages can be grouped conservatively.

SCHEMA: dict[str, Any] = {

    "type": "object",

    "properties": {

        "readable": {"type": "boolean"},

        "upright": {"type": "boolean"},

        "page_role": {"type": "string", "enum": sorted(PAGE_ROLES)},

        "uncertain_fields": {"type": "array", "items": {"type": "string"}},

        **{key: {"type": "string"} for key in FIELDS},

    },

    "required": ["readable", "upright", "page_role", "uncertain_fields", *FIELDS],

    "additionalProperties": False,

}



SYSTEM_PROMPT = """أنت نظام استخراج بيانات من صور مقررات وقرارات إدارية جزائرية.

أعد كائن JSON واحداً مطابقاً للمخطط حرفياً. اقرأ الصورة نفسها لا أي طبقة نص PDF.

قواعد أساسية:

1. استخرج فقط ما يخص القرار الحالي والشخص المعني به، وخاصة المادة الأولى/يقرر.

   لا تخلط بين القرار الحالي وبين المراسيم والقوانين والقرارات السابقة في التمهيد

   الذي يبدأ بـ بمقتضى / بناء على. ولا تُدخل أسماء الموقّعين.

2. إن لم تُذكر معلومة صراحة في هذه الصفحة ضع سلسلة فارغة ""؛ لا تخمّن مطلقاً.

   لا تستنتج الجنس من الاسم، استخرجه من السيد/السيدة قبل الاسم اذا كانت واضحة، ولا تستنتج الصنف من الرتبة.

3. حافظ على الأسماء والرتب كما هي بالعربية. لا تترجم. الأرقام والتواريخ كما تظهر.

4. decision_type: نوع الإجراء الحقيقي، مثل توظيف، تعيين، ترسيم، إعادة ترتيب，

   ترقية في الدرجة، ترقية في الرتبة، نقل، تكفل، انتداب، استيداع، استقالة، إنهاء مهام، أخرى.

   لا تعتبر ترقية في الدرجة ترقية في الرتبة. لا تقتصر على هذه الأمثلة.

5. decision_date تاريخ إصدار/توقيع المقرر الحالي (قد يكون خاتماً فرنسياً مثل 6 DEC 2015).

   decision_effective_date بدء تطبيق الإجراء (ابتداء من / يسري مفعوله / تاريخ السريان).

   grade_effective_date يخص الدرجة وحدها.

   current_rank_appointment_date تاريخ تعيين الرتبة.

   visa_date يخص تأشيرة الرقابة، وليس ختم تاريخ الإمضاء العادي. في الأغلب الغالب ختم مكتوب فيه (المراقب المالي، الرقابة المالية، المراقب الميزانياتي، المراقب المالي)

6. decision_number رقم القرار/المقرر الحالي غالباً في الرأس في الأعلى على جهة اليمين (رقم 512 أو 1391)，

   وليس أرقام القوانين، المراسيم، القرارات السابقة، رقم الاستدلال، أو رقم تأشيرة.

7. surname اللقب و given_name الاسم: افصلهما فقط إذا أمكن التمييز بينهما بوضوح؛

   وإلا اتركهما فارغين ولا تخمّن الترتيب. full_name_as_written = الاسم الكامل

   كما ورد حرفياً في المادة الأولى حتى عند عدم إمكانية فصل اللقب عن الاسم.

   استعمل البيانات الصريحة في الجسم والمادة الأولى.

8. category رقم الصنف من 1 إلى 16 فقط عندما يرد صراحة.

   grade الدرجة الفعلية (من 1 إلى 12) وليست الصنف او الرتبة، ولو كان أحدهما غير مذكور.

9. في الترقية في الرتبة: old_rank الرتبة السابقة و new_rank الرتبة الجديدة إن ذكرتا.

   في الترسيم أو التوظيف اترك هذين الحقلين فارغين إلا إذا كان هناك انتقال صريح.

   current_rank = الرتبة المنسوبة للموظف في هذا القرار حسب السياق.

10. appointment_place: مكان تعيين أو نقل الموظف وليس مكان تحرير القرار بالجزائر.

11. senior_post مثل رئيس مكتب / مكلف بالدراسات، و higher_function مثل مدير / نائب مدير.

    هذان مثالان فقط وليسا قائمة مغلقة. لا تخلط بينهما وبين رتبة الموظف.

12. employment_status: متربص / مرسم / متعاقد فقط إذا يقرر أو يذكر صراحة.

    legal_status و service_status لا تُستنتجان من المسمى المهني.

13. index_number الرقم الاستدلالي الكامل كما ورد مثل 653+98؛ لا تحسبه.

14. page_role = start إذا يبدأ قراراً مستقلاً وله عنوان/رقم أو مقدمته؛

    continuation إذا هو تتمة واضحة للصفحة السابقة من نفس القرار؛

    unknown إن تعذر التأكد. لا تستخدم عدد صفحات ملف PDF لاستنتاج ذلك.

15. readable=true إذا أمكن قراءة قدر مهم من النص، upright=true فقط إذا النص مستقيم.

16. uncertain_fields مصفوفة مفاتيح الحقول الإنجليزية التي قراءتها غير مؤكدة،

    حتى لو وضعت قيمتها مؤقتاً. notes للمعلومة الإدارية المهمة فقط، لا تكرر كامل الوثيقة.

أعد JSON فقط بلا شرح إضافي."""


CROP_PROMPT = """اقرأ هذا الجزء المكبر من صورة مقرر جزائري. أعد JSON فقط.

املأ فقط الحقول المطلوبة إذا كانت نصوصها ظاهرة بوضوح في هذا الجزء.

لا تخمّن، ولا تنقل أرقام القوانين السابقة إلى رقم القرار الحالي.

تاريخ التوقيع ليس تاريخ التأشيرة. ما لا يظهر = سلسلة فارغة."""



AR_NUMBERS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")

DIACRITICS = re.compile(r"[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06ED\u0640]")

MONTHS = {

    "جانفي": 1, "يناير": 1, "فيفري": 2, "فبراير": 2, "مارس": 3,

    "أفريل": 4, "افريل": 4, "أبريل": 4, "ابريل": 4, "ماي": 5, "مايو": 5,

    "جوان": 6, "يونيو": 6, "جويلية": 7, "جويليه": 7, "يوليو": 7,

    "أوت": 8, "اوت": 8, "غشت": 8, "أغسطس": 8, "اغسطس": 8,

    "سبتمبر": 9, "أكتوبر": 10, "اكتوبر": 10, "نوفمبر": 11, "ديسمبر": 12,

    "jan": 1, "janvier": 1, "feb": 2, "fev": 2, "fevr": 2, "février": 2,

    "fevrier": 2, "mar": 3, "mars": 3, "apr": 4, "avr": 4, "avril": 4,

    "may": 5, "mai": 5, "jun": 6, "juin": 6, "jul": 7, "juil": 7,

    "juillet": 7, "aug": 8, "aou": 8, "août": 8, "aout": 8,

    "sep": 9, "sept": 9, "septembre": 9, "oct": 10, "octobre": 10,

    "nov": 11, "novembre": 11, "dec": 12, "déc": 12, "decembre": 12,

    "décembre": 12,

}

MONTHS_N = {}  # initialized after norm()

def norm(value: Any) -> str:

    """Loose comparison only; original Arabic values are preserved in Excel."""

    text = str(value or "").translate(AR_NUMBERS)

    text = unicodedata.normalize("NFKD", text)

    text = "".join(c for c in text if not unicodedata.combining(c))

    text = DIACRITICS.sub("", text)

    text = re.sub("[إأآٱ]", "ا", text)

    text = text.replace("ى", "ي").replace("ة", "ه")

    return re.sub(r"\s+", " ", text).strip().lower()





MONTHS_N = {norm(k).strip(".،"): v for k, v in MONTHS.items()}

DATE_WORD_RE = re.compile(

    r"(?<!\d)(\d{1,2})\s*[-/.,]?\s*([\w\u0600-\u06FF]+)\.?\s*"

    r"(?:لسنه|سنه|عام|annee|an)?\s*((?:18|19|20)\d{2})(?!\d)", re.I

)

DATE_DMY_RE = re.compile(r"(?<!\d)(\d{1,2})\s*[/.-]\s*(\d{1,2})\s*[/.-]\s*((?:18|19|20)\d{2})(?!\d)")

DATE_YMD_RE = re.compile(r"(?<!\d)((?:18|19|20)\d{2})\s*[/.-]\s*(\d{1,2})\s*[/.-]\s*(\d{1,2})(?!\d)")





def parse_date(value: Any) -> dt.date | None:

    """Strict, conservative date recognition. Do not alter original text when unknown."""

    text = norm(value)

    if not text:

        return None

    for m in DATE_YMD_RE.finditer(text):

        try:

            return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))

        except ValueError:

            pass

    for m in DATE_DMY_RE.finditer(text):

        try:

            return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))

        except ValueError:

            pass

    for m in DATE_WORD_RE.finditer(text):

        name = m.group(2).strip(".،")

        month = MONTHS_N.get(name)

        if month:

            try:

                return dt.date(int(m.group(3)), month, int(m.group(1)))

            except ValueError:

                pass

    return None





def clean_value(value: Any) -> str:

    if value is None:

        return ""

    if isinstance(value, (dict, list)):

        return ""

    out = re.sub(r"\s+", " ", str(value)).strip()

    return "" if norm(out) in BLANK_MARKERS else out





def clean_model_result(raw: dict[str, Any]) -> dict[str, Any]:

    if not isinstance(raw, dict):

        raise ValueError("Model response must be a JSON object")

    fields = {key: clean_value(raw.get(key)) for key in FIELDS}

    role = raw.get("page_role", "unknown")

    unsure = raw.get("uncertain_fields", [])

    if not isinstance(unsure, list):

        unsure = []

    return {

        "fields": fields,

        "page_role": role if role in PAGE_ROLES else "unknown",

        "upright": raw.get("upright") is True,

        "readable": raw.get("readable") is True,

        "uncertain_fields": [k for k in unsure if isinstance(k, str) and k in FIELDS],

    }





@dataclass

class PageJob:

    path: Path

    page: int  # one-based page index for PDFs; 1 for an image

    page_count: int

    file_hash: str



    @property

    def label(self) -> str:

        return f"{self.path.name}#p{self.page}"





@dataclass

class PageResult:

    job: PageJob

    data: dict[str, Any]

    angle: int

    warnings: list[str] = field(default_factory=list)





@dataclass

class DecisionRecord:

    path: Path

    pages: list[int]

    fields: dict[str, str]

    warnings: list[str] = field(default_factory=list)

    uncertain: set[str] = field(default_factory=set)

    page_results: list[PageResult] = field(default_factory=list)

    duplicate_of: str = ""



    @property

    def source_label(self) -> str:

        return f"{self.path.name}#p{','.join(map(str, self.pages))}"



    @property

    def record_id(self) -> str:

        filehash = self.page_results[0].job.file_hash[:12] if self.page_results else "unknown"

        return f"{filehash}:{'-'.join(map(str, self.pages))}"





def sha256_file(path: Path) -> str:

    h = hashlib.sha256()

    with path.open("rb") as f:

        for chunk in iter(lambda: f.read(1024 * 1024), b""):

            h.update(chunk)

    return h.hexdigest()





def collect_jobs(input_dir: Path, discovery_errors: list[tuple[str, str]] | None = None) -> list[PageJob]:

    jobs: list[PageJob] = []

    for path in sorted(input_dir.iterdir(), key=lambda p: p.name.lower()):

        if not path.is_file() or path.suffix.lower() not in IMG_EXTENSIONS | {".pdf"}:

            continue

        try:

            file_hash = sha256_file(path)

            if path.suffix.lower() == ".pdf":

                with pymupdf.open(path) as pdf:

                    count = len(pdf)

                if count < 1:

                    raise ValueError("PDF contains no pages")

            else:

                count = 1

        except Exception as exc:

            if discovery_errors is None:

                raise

            discovery_errors.append((path.name, f"Could not open source file: {exc}"))

            continue

        for page in range(1, count + 1):

            jobs.append(PageJob(path, page, count, file_hash))

    return jobs





def image_for_job(job: PageJob, dpi: int) -> Image.Image:

    if job.path.suffix.lower() == ".pdf":

        with pymupdf.open(job.path) as pdf:

            pix = pdf[job.page - 1].get_pixmap(

                dpi=dpi, colorspace=pymupdf.csRGB, alpha=False, annots=True

            )

            # Everything is image-backed; never use PDF get_text() or embedded OCR text.

            return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

    with Image.open(job.path) as img:

        return ImageOps.exif_transpose(img).convert("RGB")





def encode_image(image: Image.Image, max_side: int) -> str:

    img = image.copy()

    img.thumbnail((max_side, max_side), Image.Resampling.LANCZOS)

    buffer = io.BytesIO()

    img.save(buffer, format="JPEG", quality=JPEG_QUALITY, optimize=True)

    return base64.b64encode(buffer.getvalue()).decode("ascii")





def ollama_call(model: str, image_b64: str, schema: dict[str, Any],

                prompt: str, num_ctx: int, max_tokens: int) -> dict[str, Any]:

    payload = {

        "model": model,

        "stream": False,

        "think": False,

        "keep_alive": "2m",

        "format": schema,

        "options": {"temperature": 0, "num_ctx": num_ctx, "num_predict": max_tokens},

        "messages": [

            {"role": "system", "content": prompt},

            {"role": "user", "content": "اقرأ الصورة واستخرج المعلومات حسب الحقول.",

             "images": [image_b64]},

        ],

    }

    req = urllib.request.Request(

        OLLAMA_URL + "/api/chat",

        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),

        headers={"Content-Type": "application/json"},

        method="POST",

    )

    try:

        with urllib.request.urlopen(req, timeout=TIMEOUT_SECONDS) as res:

            response = json.load(res)

    except urllib.error.HTTPError as exc:

        detail = exc.read().decode("utf-8", errors="replace")[:500]

        raise RuntimeError(f"Ollama HTTP {exc.code}: {detail}") from exc

    content = response.get("message", {}).get("content", "")

    if not isinstance(content, str):

        raise ValueError("Ollama returned no text content")

    try:

        obj = json.loads(content)

    except json.JSONDecodeError as exc:

        raise ValueError(f"Invalid JSON from Ollama: {content[:200]!r}") from exc

    if not isinstance(obj, dict):

        raise ValueError("Ollama JSON output is not an object")

    return obj





def check_ollama(model: str) -> None:

    try:

        with urllib.request.urlopen(OLLAMA_URL + "/api/tags", timeout=8) as response:

            names = {item["name"] for item in json.load(response).get("models", [])}

    except Exception as exc:

        raise RuntimeError("Ollama is not running. Start Ollama for Windows first.") from exc

    if model not in names and f"{model}:latest" not in names:

        raise RuntimeError(f"Model {model!r} is not installed. Run: ollama pull {model}")





def rotated_image(img: Image.Image, angle: int) -> Image.Image:

    return img.rotate(angle, expand=True) if angle else img





def model_page(image: Image.Image, model: str, max_side: int, num_ctx: int,

               forced_angle: int | None, detail: bool) -> tuple[dict[str, Any], int, list[str], Image.Image]:

    """Try likely orientations, and optional higher-detail crops for missing key fields."""

    angles = ([forced_angle] if forced_angle is not None else

              ([90, 270, 0, 180] if image.width > image.height else [0, 180, 90, 270]))

    possibilities = []

    for angle in angles:

        upright_image = rotated_image(image, angle)

        answer = clean_model_result(

            ollama_call(model, encode_image(upright_image, max_side), SCHEMA,

                        SYSTEM_PROMPT, num_ctx, NUM_PREDICT)

        )

        coverage = sum(bool(answer["fields"][key]) for key in KEY_FIELDS)

        score = (int(answer["upright"]) * 100 + int(answer["readable"]) * 20 + coverage)

        possibilities.append((score, angle, answer, upright_image))

        if answer["upright"] and answer["readable"] and coverage >= 2:

            break

    _, angle, answer, selected = max(possibilities, key=lambda item: item[0])

    notes: list[str] = []

    if len(possibilities) > 1:

        notes.append(f"تجربة تدوير الصفحة {len(possibilities)} مرة؛ الزاوية المختارة {angle}°")

    if not answer["upright"]:

        notes.append("اتجاه الصورة غير مؤكد")

    if not answer["readable"]:

        notes.append("الصورة غير واضحة بما يكفي")

    if detail and answer["readable"]:

        # Optional crop-assisted reading. Never overwrite a disagreeing value silently.

        crop_ranges = [

            ((0.00, 0.00, 1.00, 0.39), ["decision_number", "decision_type", "decision_subject"]),

            ((0.00, 0.39, 1.00, 0.83), ["surname", "given_name", "category", "current_rank", "grade", "decision_effective_date"]),

            ((0.00, 0.69, 1.00, 1.00), ["decision_date", "decision_effective_date", "visa_number", "visa_date"]),

        ]

        for (x0, y0, x1, y1), keys in crop_ranges:

            if all(answer["fields"].get(k) for k in keys):

                continue

            w, h = selected.size

            cropped = selected.crop((int(x0*w), int(y0*h), int(x1*w), int(y1*h)))

            crop_schema = {"type": "object", "properties": {k: {"type": "string"} for k in keys},

                           "required": keys, "additionalProperties": False}

            try:

                crop_data = ollama_call(model, encode_image(cropped, max_side), crop_schema,

                                        CROP_PROMPT + "\nالمطلوب: " + ", ".join(keys), num_ctx, 500)

                for k in keys:

                    v = clean_value(crop_data.get(k))

                    if v and not answer["fields"][k]:

                        answer["fields"][k] = v

                        answer["uncertain_fields"].append(k)

                        notes.append(f"{k}: استخرج من جزء مكبر، يحتاج تحقق")

                    elif v and norm(v) != norm(answer["fields"][k]):

                        notes.append(f"اختلاف قراءة الصورة والقصاصة في الحقل: {k}")

                        answer["uncertain_fields"].append(k)

            except Exception as exc:

                notes.append(f"فشل فحص قصاصة: {str(exc)[:90]}")

    return answer, angle, notes, selected





def cache_key(job: PageJob, model: str, dpi: int, max_side: int,

              detail: bool, angle: int | None) -> str:

    parts = (CACHE_VERSION, job.file_hash, job.page, model, dpi, max_side, detail, angle,

             NUM_CTX, NUM_PREDICT)

    return hashlib.sha256(repr(parts).encode("utf-8")).hexdigest()





def read_cache(path: Path) -> dict[str, Any] | None:

    if not path.exists():

        return None

    try:

        with path.open("r", encoding="utf-8") as fp:

            data = json.load(fp)

        if isinstance(data.get("data", {}).get("fields"), dict):

            return data

    except (OSError, json.JSONDecodeError, AttributeError):

        pass

    return None





def atomic_json(path: Path, data: dict[str, Any]) -> None:

    path.parent.mkdir(parents=True, exist_ok=True)

    temporary = path.with_name(path.name + ".tmp")

    with temporary.open("w", encoding="utf-8") as fp:

        json.dump(data, fp, ensure_ascii=False, indent=2)

    temporary.replace(path)





def validate_record(record: DecisionRecord) -> list[str]:

    f = record.fields

    warnings = list(record.warnings)

    if not f["decision_number"]:

        warnings.append("رقم المقرر غير مستخرج")

    if not f["decision_date"]:

        warnings.append("تاريخ المقرر غير مستخرج")

    if not f["surname"] or not f["given_name"]:

        warnings.append("الاسم واللقب غير مكتملين")

    if not f["decision_type"]:

        warnings.append("نوع المقرر غير محدد")

    if f["category"]:

        cat = norm(f["category"]).lstrip("0") or "0"

        if not cat.isdecimal() or not 1 <= int(cat) <= 16:

            warnings.append("الصنف خارج 1..16 أو غير رقمي؛ راجع المصدر")

    if f["gender"] and norm(f["gender"]) not in ("ذكر", "انثي", "انثى"):

        warnings.append("الجنس غير مطابق للقيم المتوقعة، لا يُصحح آلياً")

    for key in DATE_FIELDS:

        value = f[key]

        if value and parse_date(value) is None:

            warnings.append(f"تاريخ/صيغة غير محددة آلياً في {FIELDS[key]}")

    if "ترقية في الرتبة" in norm(f["decision_type"]) and not f["new_rank"]:

        warnings.append("ترقية في الرتبة لكن الرتبة الجديدة غير مستخرجة")

    if record.uncertain:

        warnings.append("حقول غير مؤكدة: " + ", ".join(FIELDS[k] for k in sorted(record.uncertain)))

    return list(dict.fromkeys(warnings))





def can_merge(a: DecisionRecord, page: PageResult) -> bool:

    # Only one source file can be merged, and contradictory identities block merges.

    if a.path != page.job.path:

        return False

    newer = page.data["fields"]

    older = a.fields

    for key in IDENTITY_FIELDS:

        if older[key] and newer[key]:

            if key in DATE_FIELDS:

                d0, d1 = parse_date(older[key]), parse_date(newer[key])

                if d0 and d1 and d0 != d1:

                    return False

            elif norm(older[key]) != norm(newer[key]):

                return False

    return True





def start_record(page: PageResult) -> DecisionRecord:

    return DecisionRecord(

        path=page.job.path, pages=[page.job.page],

        fields=dict(page.data["fields"]),

        uncertain=set(page.data["uncertain_fields"]),

        warnings=list(page.warnings), page_results=[page],

    )





def merge_page(record: DecisionRecord, page: PageResult, automatic: bool) -> None:

    record.pages.append(page.job.page)

    record.page_results.append(page)

    record.warnings.extend(page.warnings)

    record.uncertain.update(page.data["uncertain_fields"])

    for key in FIELDS:

        old, new = record.fields[key], page.data["fields"][key]

        if not old and new:

            record.fields[key] = new

        elif old and new and norm(old) != norm(new):

            if key in DATE_FIELDS and parse_date(old) and parse_date(old) == parse_date(new):

                continue

            if key == "notes":

                record.fields[key] = old + " | " + new

            else:

                record.warnings.append(f"تعارض في {FIELDS[key]} بين صفحات القرار؛ احتُفظ بالقيمة الأولى")

                record.uncertain.add(key)

    if automatic:

        record.warnings.append("تم تجميع الصفحات آلياً؛ راجع أنها تتبع نفس المقرر")





def parse_group_config(path: Path) -> dict[str, list[list[int]]]:

    if not path.exists():

        return {}

    with path.open("r", encoding="utf-8-sig") as fp:

        raw = json.load(fp)

    if not isinstance(raw, dict):

        raise ValueError("page_groups.json must be an object mapping PDF filenames to page groups")

    result = {}

    for name, groups in raw.items():

        if not isinstance(groups, list) or any(not isinstance(g, list) or not g for g in groups):

            raise ValueError(f"Invalid group list for {name!r}")

        if any(g != sorted(g) for g in groups):

            raise ValueError(f"Pages in every manual group must be ordered: {name!r}")

        page_list = [p for g in groups for p in g]

        if any(type(p) is not int or p < 1 for p in page_list):

            raise ValueError(f"Page numbers must be positive integers: {name!r}")

        if len(page_list) != len(set(page_list)):

            raise ValueError(f"Repeated page numbers in groups: {name!r}")

        result[name] = groups

    return result





def group_pages(pages: list[PageResult], manual: dict[str, list[list[int]]]) -> list[DecisionRecord]:

    by_file: dict[Path, list[PageResult]] = defaultdict(list)

    for result in pages:

        by_file[result.job.path].append(result)

    records = []

    for path, page_list in by_file.items():

        page_list.sort(key=lambda pr: pr.job.page)

        incomplete_file = len(page_list) < page_list[0].job.page_count

        by_number = {p.job.page: p for p in page_list}

        if path.name in manual:

            planned = manual[path.name]

            expected = set(range(1, page_list[0].job.page_count + 1))

            provided = {p for g in planned for p in g}

            if expected != provided:

                raise ValueError(f"Manual page groups for {path.name} must cover every page exactly once")

            for group in planned:

                present = [by_number[p] for p in group if p in by_number]

                if not present:

                    continue  # limit/cached/incomplete run

                rec = start_record(present[0])

                for page in present[1:]:

                    merge_page(rec, page, automatic=False)

                if len(present) != len(group):

                    rec.warnings.append("مجموعة صفحات غير مكتملة (تشغيل تجريبي --limit أو أخطاء)")

                elif incomplete_file:

                    rec.warnings.append("لم تتم معالجة كل صفحات PDF؛ راجع أي قرارات ناقصة")

                records.append(rec)

            continue

        last: DecisionRecord | None = None

        for p in page_list:

            role = p.data["page_role"]

            # Check adjacency; don't jump across failed/missing pages.

            adjacent = last is not None and p.job.page == last.pages[-1] + 1

            if role == "continuation" and adjacent and last and can_merge(last, p):

                merge_page(last, p, automatic=True)

            else:

                if role == "continuation":

                    p.warnings.append("صفحة تتمة بلا بداية مطابقة؛ لم تُدمج")

                elif role == "unknown" and p.job.page_count > 1:

                    p.warnings.append("بداية/تتمة الصفحة غير مؤكدة؛ لم تُدمج تلقائياً")

                last = start_record(p)

                records.append(last)

            if incomplete_file and last is not None:

                last.warnings.append("لم تتم معالجة كل صفحات PDF؛ التشغيل الحالي جزئي")

    return records





def mark_duplicates(records: list[DecisionRecord]) -> None:

    seen: dict[tuple[str, ...], str] = {}

    same_file: dict[tuple[str, tuple[int, ...]], str] = {}

    for rec in records:

        f = rec.fields

        key = tuple(norm(f[k]) if k not in DATE_FIELDS else

                    (parse_date(f[k]).isoformat() if parse_date(f[k]) else norm(f[k]))

                    for k in IDENTITY_FIELDS)

        file_key = (rec.page_results[0].job.file_hash, tuple(rec.pages))

        if file_key in same_file:

            rec.duplicate_of = same_file[file_key]

        else:

            same_file[file_key] = rec.record_id

        # Match when all identity parts exist, or fall back to a literal full name.

        # This only flags a potential duplicate; no decision is discarded.

        if all(key):

            identity = key

        elif all((f["decision_number"], f["decision_date"],

                  f["full_name_as_written"], f["decision_type"])):

            identity = (key[0], key[1], norm(f["full_name_as_written"]), key[4])

        else:

            identity = ()

        if identity:

            if identity in seen:

                rec.duplicate_of = seen[identity]

            else:

                seen[identity] = rec.record_id

        if rec.duplicate_of:

            rec.warnings.append(f"مقرر مشابه/مكرر محتمَل: {rec.duplicate_of}؛ تحقق قبل حذف النسخة")





def load_corrections(path: Path) -> dict[str, dict[str, str]]:

    """Optional durable human corrections, keyed by the stable source record_id."""

    result: dict[str, dict[str, str]] = defaultdict(dict)

    if not path.exists():

        return result

    with path.open("r", encoding="utf-8-sig", newline="") as fp:

        reader = csv.DictReader(fp)

        if not {"record_id", "field", "value"}.issubset(reader.fieldnames or []):

            raise ValueError("corrections.csv must contain record_id,field,value")

        for row in reader:

            record_id = (row.get("record_id") or "").strip()

            key = (row.get("field") or "").strip()

            if key in FIELDS.values():

                key = next(k for k, v in FIELDS.items() if v == key)

            if record_id and key in FIELDS:

                result[record_id][key] = row.get("value") or ""

            elif record_id or key:

                raise ValueError(f"Invalid correction record_id={record_id!r} field={key!r}")

    return result





def apply_corrections(records: list[DecisionRecord], overrides: dict[str, dict[str, str]]) -> None:

    for rec in records:

        if rec.record_id not in overrides:

            continue

        for key, val in overrides[rec.record_id].items():

            rec.fields[key] = clean_value(val)  # empty intentionally clears a field

            rec.uncertain.discard(key)

        rec.warnings.append("تم تطبيق تصحيحات بشرية محفوظة من corrections.csv")





def excel_cell(key: str, text: str) -> Any:

    # Keep the original text if not unambiguously parseable.

    if key in DATE_FIELDS:

        return parse_date(text) or safe_excel_text(text)

    # Prevent untrusted document text from becoming an Excel formula.

    return safe_excel_text(text)





def safe_excel_text(value: str) -> str:

    if not value:

        return ""

    if value[0] in ("=", "+", "-", "@"):

        return "'" + value

    return value





def export_excel(records: list[DecisionRecord], errors: list[tuple[str, str]], outdir: Path) -> Path:

    outdir.mkdir(parents=True, exist_ok=True)

    filename = "المقررات_" + dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".xlsx"

    output = outdir / filename

    wb = Workbook()

    ws = wb.active

    ws.title = "المقررات"

    extras = ["معرّف السجل", "ملف المصدر", "الصفحات", "حالة المراجعة",

              "أسباب المراجعة", "مكرر محتمل من", "حقول غير مؤكدة"]

    headers = [*FIELDS.values(), *extras]

    ws.append(headers)

    ws.sheet_view.rightToLeft = True

    ws.freeze_panes = "A2"

    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(len(records) + 1, 2)}"



    for rec in records:

        reasons = validate_record(rec)

        status = "يحتاج مراجعة" if reasons else "غير مُراجَع - تحقق قبل الاعتماد"

        ws.append([

            *(excel_cell(k, rec.fields[k]) for k in FIELDS),

            rec.record_id, safe_excel_text(rec.path.name), ", ".join(map(str, rec.pages)), status,

            safe_excel_text(" | ".join(reasons)), rec.duplicate_of,

            ", ".join(FIELDS[k] for k in sorted(rec.uncertain)),

        ])

        row = ws.max_row

        for ci, key in enumerate(FIELDS, 1):

            cell = ws.cell(row, ci)

            cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=False)

            if key in DATE_FIELDS and isinstance(cell.value, dt.date):

                cell.number_format = "yyyy-mm-dd"

        ws.cell(row, len(FIELDS) + 4).fill = PatternFill(

            "solid", fgColor="FFF2CC" if reasons else "EAF2F8"

        )



    review = wb.create_sheet("للمراجعة")

    review.sheet_view.rightToLeft = True

    review.append(["معرّف السجل", "المصدر", "نوع المقرر", "الاسم واللقب", "سبب المراجعة"])

    for rec in records:

        reasons = validate_record(rec)

        if reasons:

            review.append([rec.record_id, safe_excel_text(rec.source_label), safe_excel_text(rec.fields["decision_type"]),

                           safe_excel_text((rec.fields["surname"] + " " + rec.fields["given_name"]).strip()),

                           safe_excel_text(" | ".join(reasons))])



    fail = wb.create_sheet("أخطاء المعالجة")

    fail.sheet_view.rightToLeft = True

    fail.append(["الملف/الصفحة", "الخطأ", "إجراء مقترح"])

    for label, message in errors:

        fail.append([safe_excel_text(label), safe_excel_text(str(message)[:1000]), "أعد تشغيل السكربت؛ الصفحات الفاشلة لا تُخزن في الكاش"])



    for sh in wb:

        sh.row_dimensions[1].height = 32

        for cell in sh[1]:

            cell.fill = PatternFill("solid", fgColor="153B5C")

            cell.font = Font(bold=True, color="FFFFFF", size=11)

            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        sh.freeze_panes = "A2"

        sh.sheet_view.rightToLeft = True

        if sh == ws:

            for j, label in enumerate(headers, 1):

                sh.column_dimensions[get_column_letter(j)].width = min(38, max(14, len(label) + 5))

            for j in (len(FIELDS) + 2, len(FIELDS) + 5):

                sh.column_dimensions[get_column_letter(j)].width = 40

        else:

            for j in range(1, sh.max_column + 1):

                sh.column_dimensions[get_column_letter(j)].width = 30 if j < sh.max_column else 85

    wb.save(output)

    return output





def load_json_if_exists(path: Path) -> dict[str, Any]:

    if not path.exists():

        return {}

    with path.open("r", encoding="utf-8-sig") as fp:

        data = json.load(fp)

    if not isinstance(data, dict):

        raise ValueError(f"{path} must contain a JSON object")

    return data





def progress_bar(done: float, total: int, width: int = 24) -> str:
    """Visual progress; running-page fractions are estimates, not model telemetry."""
    fraction = max(0.0, min(1.0, done / total)) if total else 1.0
    filled = round(width * fraction)
    return f"[{'#' * filled}{'-' * (width - filled)}] {fraction * 100:5.1f}%"


def estimated_page_fraction(elapsed: float, expected_seconds: float) -> float:
    """Time-based estimate that approaches 95%, never 100%, until OCR finishes.

    The estimate adapts to previous successfully processed pages. It is NOT a
    measurement of how far Ollama has got inside a single OCR request.
    """
    import math
    expected_seconds = max(1.0, expected_seconds)
    return min(0.95, 0.95 * (1.0 - math.exp(-2.0 * elapsed / expected_seconds)))


def run_with_estimated_progress(job: PageJob, callback, file_done: int,
                                file_total: int, overall_done: int,
                                overall_total: int, expected_seconds: float,
                                update_seconds: float = 1.0):
    """Show estimated live progress during the blocking OCR operation."""
    stopped = threading.Event()
    started = time.monotonic()

    def report():
        # Print full lines for predictable output in VS Code/PowerShell logs.
        while not stopped.wait(update_seconds):
            elapsed = time.monotonic() - started
            page_fraction = estimated_page_fraction(elapsed, expected_seconds)
            file_fraction = file_done + page_fraction
            overall_fraction = overall_done + page_fraction
            print(
                f"  ESTIMATED | {job.path.name} page {job.page}/{job.page_count} "
                f"| FILE {progress_bar(file_fraction, file_total)} "
                f"| ALL {progress_bar(overall_fraction, overall_total)} "
                f"| {elapsed:.0f}s elapsed (not exact)",
                flush=True,
            )

    worker = threading.Thread(target=report, daemon=True)
    worker.start()
    try:
        return callback()
    finally:
        stopped.set()
        worker.join(timeout=1.0)


def process(args: argparse.Namespace) -> tuple[list[DecisionRecord], list[tuple[str, str]], Path]:

    errors: list[tuple[str, str]] = []

    jobs = collect_jobs(args.input, errors)

    if not jobs:

        if errors:

            return [], errors, export_excel([], errors, args.output)

        raise RuntimeError(f"No PDFs/images found in {args.input.resolve()}")

    page_groups = parse_group_config(args.groups)

    rotate_map = load_json_if_exists(args.rotations)

    corrections = load_corrections(args.corrections)

    pages: list[PageResult] = []

    uncached_done = 0

    live_model_checked = False

    # Each source PDF/image has its own page counter, independent of other files.
    file_totals = {job.path: job.page_count for job in jobs}
    file_finished = {path: 0 for path in file_totals}
    total_finished = 0
    total_pages = len(jobs)
    observed_ocr_seconds: list[float] = []

    def show_progress(job: PageJob, state: str) -> None:
        nonlocal total_finished
        file_finished[job.path] += 1
        total_finished += 1
        current = file_finished[job.path]
        total = file_totals[job.path]
        print(f"  FILE {job.path.name}: {current}/{total} pages "
              f"{progress_bar(current, total)} ({state})", flush=True)
        print(f"  ALL FILES: {total_finished}/{total_pages} pages "
              f"{progress_bar(total_finished, total_pages)}", flush=True)

    print(f"Starting {len(file_totals)} file(s), {total_pages} page(s).", flush=True)
    for ix, job in enumerate(jobs, 1):

        forced_angle = rotate_map.get(job.label, rotate_map.get(job.path.name))

        if forced_angle is not None:

            if type(forced_angle) is not int or forced_angle not in (0, 90, 180, 270):

                raise ValueError(f"Rotation must be 0/90/180/270 for {job.label}")

        cachefile = args.cache / (cache_key(job, args.model, args.pdf_dpi, args.max_side,

                                            args.detail, forced_angle) + ".json")

        cached = None if args.force else read_cache(cachefile)

        if cached:

            pages.append(PageResult(job, cached["data"], cached["angle"], cached.get("warnings", [])))

            print(f"[{ix}/{len(jobs)}] cached {job.label}", flush=True)
            show_progress(job, "cached")
            continue

        if args.limit is not None and uncached_done >= args.limit:

            continue

        if not live_model_checked:

            check_ollama(args.model)  # Fail early instead of creating an empty workbook.

            live_model_checked = True

        try:

            print(f"[{ix}/{len(jobs)}] reading {job.label} ...", flush=True)
            print(f"  FILE {job.path.name}: {file_finished[job.path]}/{file_totals[job.path]} pages "
                  f"{progress_bar(file_finished[job.path], file_totals[job.path])} "
                  "(estimated live progress starts below)", flush=True)
            start = time.monotonic()

            with image_for_job(job, args.pdf_dpi) as img:

                # Historical durations improve the estimate as pages finish.
                # The first page starts with a configurable 90-second guess.
                expected_seconds = (sum(observed_ocr_seconds[-5:]) /
                                    len(observed_ocr_seconds[-5:])) if observed_ocr_seconds else 90.0
                ocr_started = time.monotonic()
                result, angle, notes, selected = run_with_estimated_progress(
                    job,
                    lambda: model_page(
                        img, args.model, args.max_side, args.num_ctx,
                        forced_angle, args.detail
                    ),
                    file_finished[job.path], file_totals[job.path],
                    total_finished, total_pages, expected_seconds,
                )
                observed_ocr_seconds.append(time.monotonic() - ocr_started)

                args.debug.mkdir(parents=True, exist_ok=True)

                thumbnail = selected.copy()

                thumbnail.thumbnail((args.max_side, args.max_side))

                safe_name = hashlib.sha1(job.label.encode("utf-8")).hexdigest()[:12]

                thumbnail.save(args.debug / f"{safe_name}_p{job.page}.jpg", quality=85)

            if not result["readable"]:

                notes.append("تعذر تأكيد قراءة الصفحة، لا تعتمد الحقول دون مراجعة")

            pages.append(PageResult(job, result, angle, notes))

            atomic_json(cachefile, {"data": result, "angle": angle, "warnings": notes,

                                    "source": job.label, "model": args.model})

            uncached_done += 1

            print(f"    saved | {time.monotonic() - start:.1f}s | rotation={angle}°", flush=True)
            show_progress(job, "saved")

        except (Exception, KeyboardInterrupt) as exc:

            if isinstance(exc, KeyboardInterrupt):

                print("\nInterrupted; successful pages are safely cached.")

                break

            errors.append((job.label, str(exc)))

            print(f"    ERROR {job.label}: {exc}", file=sys.stderr, flush=True)
            show_progress(job, "ERROR - review needed")

    if total_finished < total_pages:
        print(f"Partial run: {total_finished}/{total_pages} pages attempted "
              "(--limit or interrupted).", flush=True)
    records = group_pages(pages, page_groups)

    # Correct before duplicate matching so corrected decision IDs can be compared.

    apply_corrections(records, corrections)

    mark_duplicates(records)

    output = export_excel(records, errors, args.output)

    return records, errors, output





def arguments(argv: list[str] | None = None) -> argparse.Namespace:

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)

    parser.add_argument("--input", type=Path, default=Path("decisions"), help="Folder of scanned PDF/image files")

    parser.add_argument("--output", type=Path, default=Path("exports"), help="Folder for Excel exports")

    parser.add_argument("--cache", type=Path, default=Path(".cache"), help="Successful page-extraction cache")

    parser.add_argument("--debug", type=Path, default=Path("debug_pages"), help="Page previews sent to the model")

    parser.add_argument("--groups", type=Path, default=Path("page_groups.json"), help="Optional manual PDF page groupings")

    parser.add_argument("--rotations", type=Path, default=Path("rotations.json"), help="Optional manual rotation angles")

    parser.add_argument("--corrections", type=Path, default=Path("corrections.csv"), help="Optional corrections persisted between runs")

    parser.add_argument("--model", default=DEFAULT_MODEL)

    parser.add_argument("--pdf-dpi", type=int, default=PDF_DPI)

    parser.add_argument("--max-side", type=int, default=MAX_SIDE)

    parser.add_argument("--num-ctx", type=int, default=NUM_CTX)

    parser.add_argument("--limit", type=int, help="Process up to N *uncached* pages this run")

    parser.add_argument("--force", action="store_true", help="Ignore successful cache and reread pages")

    parser.add_argument("--detail", action="store_true", help="Slow extra passes on image crops")

    args = parser.parse_args(argv)

    if args.pdf_dpi < 72 or args.max_side < 600 or args.num_ctx < 1024:

        parser.error("--pdf-dpi must be >=72, --max-side >=600, and --num-ctx >=1024")

    if args.limit is not None and args.limit < 1:

        parser.error("--limit must be a positive integer")

    return args





def main(argv: list[str] | None = None) -> int:

    args = arguments(argv)

    try:

        records, errors, output = process(args)

    except (OSError, ValueError, RuntimeError) as exc:

        print(f"FATAL ERROR: {exc}", file=sys.stderr)

        return 1

    print(f"\nExcel export: {output.resolve()}")

    print(f"Records: {len(records)} | errors: {len(errors)}")

    print("ALL records require human verification before use in official employee files.")

    return 0 if not errors else 2





if __name__ == "__main__":

    sys.exit(main())
