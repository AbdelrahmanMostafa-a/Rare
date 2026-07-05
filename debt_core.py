"""
debt_core.py
منطق قراءة وتنظيف بيانات المديونية (من Excel / TXT / صورة) وملفات PDF ومطابقتها.
منفصل عن واجهة Streamlit عشان يسهل اختباره وصيانته.
"""

import re

import pandas as pd
import pdfplumber

REQUIRED_COLS = ["رقم العميل", "اسم العميل", "صافي المديونيه"]

# عمود اختياري: لو موجود في ملف المديونية، بنستخدمه لتحديد العملاء
# اللي "يجب المرور عليهم" (تجاوزوا المدة المطلوب سدادها)
OVERDUE_COL = "اجمالي تجاوز المده المطلوب سداده"


def _clean_cell(v):
    v = "" if v is None else str(v).strip()
    return "" if v.lower() == "nan" else v


def _normalize_ar(s):
    # بيوحّد الفرق الشائع بين "ة" و"ه" ويشيل الفراغات الزيادة، عشان مطابقة
    # اسم العمود تفضل شغالة لو اختلف الإملاء شوية بين نسخ الملف
    s = _clean_cell(s).replace("ة", "ه")
    return re.sub(r"\s+", " ", s)


def _find_overdue_column(columns):
    target = _normalize_ar(OVERDUE_COL)
    for col in columns:
        if _normalize_ar(col) == target:
            return col
    return None


def find_header_row(rows):
    """يرجع (index, header) لأول صف فيه كل الأعمدة المطلوبة - في أي ترتيب وأي مكان."""
    for i, row in enumerate(rows):
        cleaned = [_clean_cell(c) for c in row]
        if all(col in cleaned for col in REQUIRED_COLS):
            return i, cleaned
    return None, None


def rows_to_dataframe(rows, header_row):
    ncols = len(header_row)
    fixed_rows = []
    for row in rows:
        row = list(row) + [""] * (ncols - len(row))
        fixed_rows.append(row[:ncols])
    return pd.DataFrame(fixed_rows, columns=header_row)


def parse_number(v):
    s = _clean_cell(v)
    if s == "":
        return None
    s = s.replace(" ", "")
    if re.match(r"^-?\d+,\d{1,2}$", s):
        # فاصلة كفاصل عشري (مثل 1500,00)
        s = s.replace(",", ".")
    else:
        # فاصلة كفاصل آلاف (مثل 1,500.00)
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def normalize_debt(raw_debt):
    missing = [c for c in REQUIRED_COLS if c not in raw_debt.columns]
    if missing:
        raise ValueError("الأعمدة الناقصة: " + "، ".join(missing))

    overdue_actual_col = _find_overdue_column(raw_debt.columns)

    cols_to_take = list(REQUIRED_COLS)
    if overdue_actual_col:
        cols_to_take.append(overdue_actual_col)

    debt = raw_debt[cols_to_take].copy()
    if overdue_actual_col:
        debt = debt.rename(columns={overdue_actual_col: OVERDUE_COL})

    debt["رقم العميل"] = debt["رقم العميل"].map(_clean_cell)
    debt = debt[debt["رقم العميل"] != ""]
    debt["رقم العميل"] = debt["رقم العميل"].str.replace(r"\.0$", "", regex=True)

    debt["اسم العميل"] = debt["اسم العميل"].map(_clean_cell)
    debt["صافي المديونيه"] = debt["صافي المديونيه"].map(parse_number)

    if overdue_actual_col:
        debt[OVERDUE_COL] = debt[OVERDUE_COL].map(parse_number)

    duplicate_count = int(debt["رقم العميل"].duplicated().sum())
    return debt.reset_index(drop=True), duplicate_count


def read_excel_debt(file):
    xls = pd.ExcelFile(file)
    preferred = "مديونيه المباشر"
    sheets = xls.sheet_names
    ordered = ([preferred] if preferred in sheets else []) + [s for s in sheets if s != preferred]

    for sheet in ordered:
        raw = xls.parse(sheet_name=sheet, header=None)
        rows = raw.values.tolist()
        header_idx, header_row = find_header_row(rows)
        if header_idx is not None:
            return rows_to_dataframe(rows[header_idx + 1:], header_row)

    raise ValueError(
        "تعذر العثور على الأعمدة المطلوبة (رقم العميل / اسم العميل / صافي المديونيه) في أي شيت بالملف"
    )


def _decode_bytes(b):
    for enc in ("utf-8-sig", "utf-8", "cp1256"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            continue
    return b.decode("utf-8", errors="replace")


def read_txt_debt(file):
    content = file.read()
    if isinstance(content, bytes):
        content = _decode_bytes(content)

    lines = [l for l in content.splitlines() if l.strip() != ""]

    for sep in ["\t", ",", ";", "|"]:
        rows = [line.split(sep) for line in lines]
        header_idx, header_row = find_header_row(rows)
        if header_idx is not None:
            return rows_to_dataframe(rows[header_idx + 1:], header_row)

    rows = [re.split(r"\s{2,}", line.strip()) for line in lines]
    header_idx, header_row = find_header_row(rows)
    if header_idx is not None:
        return rows_to_dataframe(rows[header_idx + 1:], header_row)

    raise ValueError("تعذر التعرف على شكل الأعمدة داخل ملف TXT")


def read_image_debt(file):
    try:
        import pytesseract
        from PIL import Image
    except ImportError as e:
        raise ValueError(
            "قراءة الصور تحتاج تثبيت مكتبتي pytesseract و Pillow، "
            "وبرنامج tesseract-ocr (مع حزمة اللغة العربية) على السيرفر"
        ) from e

    image = Image.open(file)
    text = pytesseract.image_to_string(image, lang="ara+eng")
    lines = [l for l in text.splitlines() if l.strip() != ""]
    rows = [re.split(r"\s{2,}", line.strip()) for line in lines]

    header_idx, header_row = find_header_row(rows)
    if header_idx is None:
        raise ValueError(
            "تعذر التعرف على الأعمدة المطلوبة من الصورة - جرب صورة أوضح/أقرب، "
            "أو تأكد إن حزمة اللغة العربية لـ tesseract متاحة على السيرفر"
        )
    return rows_to_dataframe(rows[header_idx + 1:], header_row)


def read_debt_file(file):
    name = file.name.lower()

    if name.endswith((".xlsx", ".xls")):
        raw_debt = read_excel_debt(file)
    elif name.endswith(".txt"):
        raw_debt = read_txt_debt(file)
    elif name.endswith((".png", ".jpg", ".jpeg")):
        raw_debt = read_image_debt(file)
    else:
        raise ValueError("صيغة الملف غير مدعومة")

    return normalize_debt(raw_debt)


def read_pdf(file):
    rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                client = re.search(r"(470\d{7,8})", line)
                if not client:
                    continue
                nums = re.findall(r"\d+,\d{2}", line)
                if len(nums) >= 3:
                    rows.append({
                        "رقم العميل": client.group(1),
                        "رصيد PDF": float(nums[0].replace(",", ".")),
                    })
    return pd.DataFrame(rows)


def compute_status(r):
    if pd.isna(r["صافي المديونيه"]):
        return "PDF فقط"
    if pd.isna(r["رصيد PDF"]):
        return "مديونية فقط"
    if abs(r["الفرق"]) < 1:
        return "مطابق"
    return "يوجد فرق"


FOLLOWUP_SHEET_NAME = "عملاء يجب المرور عليهم"


def build_followup_list(result):
    """
    قائمة العملاء اللي عندهم قيمة أكبر من صفر في عمود "اجمالي تجاوز المده
    المطلوب سداده" (يعني تجاوزوا المدة المتاحة للسداد)، مرتبة حسب الرصيد
    الموجود في PDF من الأكبر للأصغر (العملاء اللي مفيش لهم رصيد PDF بييجوا
    في الآخر). ترجع None لو العمود ده غير موجود في الملف الأصلي.
    """
    if OVERDUE_COL not in result.columns:
        return None

    followup = (
        result[result[OVERDUE_COL].fillna(0) > 0]
        .sort_values("رصيد PDF", ascending=False, na_position="last")
        .reset_index(drop=True)
    )
    followup.insert(0, "الترتيب", range(1, len(followup) + 1))
    return followup
