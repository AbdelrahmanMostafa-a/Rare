"""
debt_core.py
منطق قراءة وتنظيف بيانات المديونية (من Excel / TXT / صورة) وملفات PDF ومطابقتها.
منفصل عن واجهة Streamlit عشان يسهل اختباره وصيانته.
"""

import re
import unicodedata

import pandas as pd
import pdfplumber

REQUIRED_COLS = ["رقم العميل", "اسم العميل", "صافي المديونيه"]

# عمود اختياري: لو موجود في ملف المديونية، بنستخدمه لتحديد العملاء
# اللي "يجب المرور عليهم" (تجاوزوا المدة المطلوب سدادها)
OVERDUE_COL = "اجمالي تجاوز المده المطلوب سداده"

# عمود اختياري: "الخط" (المنطقة/خط التوزيع) - بنستخدمه لمعرفة أي خطوط
# اتغطت فعلاً بملفات الـPDF المرسلة، عشان نفرّق "مديونية فقط" الحقيقية عن
# عميل خطه لسه مبعتش PDF أصلاً
ROUTE_COL = "الخط"

NOT_SENT_STATUS = "الخط لم يُرسل"


def _clean_cell(v):
    v = "" if v is None else str(v).strip()
    return "" if v.lower() == "nan" else v


def _normalize_ar(s):
    # بيوحّد الفرق الشائع بين "ة" و"ه" ويشيل الفراغات الزيادة، عشان مطابقة
    # اسم العمود تفضل شغالة لو اختلف الإملاء شوية بين نسخ الملف
    s = _clean_cell(s).replace("ة", "ه")
    return re.sub(r"\s+", " ", s)


def _find_column(columns, target):
    normalized_target = _normalize_ar(target)
    for col in columns:
        if _normalize_ar(col) == normalized_target:
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

    overdue_actual_col = _find_column(raw_debt.columns, OVERDUE_COL)
    route_actual_col = _find_column(raw_debt.columns, ROUTE_COL)

    cols_to_take = list(REQUIRED_COLS)
    if overdue_actual_col:
        cols_to_take.append(overdue_actual_col)
    if route_actual_col:
        cols_to_take.append(route_actual_col)

    debt = raw_debt[cols_to_take].copy()
    rename_map = {}
    if overdue_actual_col:
        rename_map[overdue_actual_col] = OVERDUE_COL
    if route_actual_col:
        rename_map[route_actual_col] = ROUTE_COL
    if rename_map:
        debt = debt.rename(columns=rename_map)

    debt["رقم العميل"] = debt["رقم العميل"].map(_clean_cell)
    debt = debt[debt["رقم العميل"] != ""]
    debt["رقم العميل"] = debt["رقم العميل"].str.replace(r"\.0$", "", regex=True)

    debt["اسم العميل"] = debt["اسم العميل"].map(_clean_cell)
    debt["صافي المديونيه"] = debt["صافي المديونيه"].map(parse_number)

    if overdue_actual_col:
        debt[OVERDUE_COL] = debt[OVERDUE_COL].map(parse_number)
    if route_actual_col:
        debt[ROUTE_COL] = debt[ROUTE_COL].map(_clean_cell)

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


def _readable_arabic(text):
    """
    ملفات PDF بعض برامج التقارير (زي Konnash) بتحفظ النص بشكل معكوس
    وبحروف عرض (presentation forms)، فـ pdfplumber بيطلعه مقلوب وغير
    قابل للمطابقة المباشرة. عكس النص كامل ثم تطبيع NFKC بيرجعه نص عربي
    عادي قابل للمقارنة.
    """
    return unicodedata.normalize("NFKC", text[::-1])


def _detect_pdf_report_type(first_page_text):
    readable = _readable_arabic(first_page_text)
    if "حسب العميل" in readable:
        return "عميل"
    if "حسب الكليان" in readable:
        return "كليان"
    if "حسب الفورنيسور" in readable:
        return "فورنيسور"
    return None


VALID_CUSTOMER_REPORT_TYPES = {"عميل", "كليان", "فورنيسور"}


def read_pdf(file):
    rows = []
    with pdfplumber.open(file) as pdf:
        first_page_text = pdf.pages[0].extract_text() or ""
        report_type = _detect_pdf_report_type(first_page_text)

        if report_type is not None and report_type not in VALID_CUSTOMER_REPORT_TYPES:
            raise ValueError(
                f"هذا تقرير \"الوضع حسب {report_type}\" مش تقرير عملاء - تم تجاهله"
            )

        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                client = re.search(r"(470\d{7,8})", line)
                if not client:
                    continue
                nums = re.findall(r"\d+,\d{2}", line)
                if not nums:
                    continue
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
    قائمة العملاء اللي (أ) موجودين في ملفات PDF المرسلة و(ب) عندهم قيمة أكبر
    من صفر في عمود "اجمالي تجاوز المده المطلوب سداده"، مرتبة حسب الرصيد
    الموجود في PDF من الأكبر للأصغر. ترجع None لو العمود ده غير موجود
    في ملف المديونية الأصلي.
    """
    if OVERDUE_COL not in result.columns:
        return None

    in_pdf = result["رصيد PDF"].notna()
    has_overdue = result[OVERDUE_COL].fillna(0) > 0

    followup = (
        result[in_pdf & has_overdue]
        .sort_values("رصيد PDF", ascending=False)
        .reset_index(drop=True)
    )
    followup.insert(0, "الترتيب", range(1, len(followup) + 1))
    return followup


def compute_route_coverage(result):
    """
    لكل "خط": عدد عملائه في شيت المديونية، وعدد اللي ظهروا فعلاً في ملفات
    PDF المرسلة، ونسبة التغطية. مرتبة من الأعلى تغطية للأقل. ترجع None لو
    عمود "الخط" غير متاح في ملف المديونية.
    """
    if ROUTE_COL not in result.columns:
        return None

    total_by_route = result.groupby(ROUTE_COL)["رقم العميل"].count()
    covered_by_route = (
        result[result["رصيد PDF"].notna()].groupby(ROUTE_COL)["رقم العميل"].count()
    )

    coverage = pd.DataFrame({
        ROUTE_COL: total_by_route.index,
        "عدد عملاء الخط": total_by_route.values,
    })
    coverage["عدد الموجودين في PDF"] = (
        coverage[ROUTE_COL].map(covered_by_route).fillna(0).astype(int)
    )
    coverage["نسبة التغطية %"] = (
        coverage["عدد الموجودين في PDF"] / coverage["عدد عملاء الخط"] * 100
    ).round(1)

    return coverage.sort_values("نسبة التغطية %", ascending=False).reset_index(drop=True)


def refine_debt_only_by_route(result):
    """
    لو عمود "الخط" متاح: بتضيف عمود "نسبة تغطية الخط %" لكل صف، وتقسّم
    حالة "مديونية فقط" لقسمين - "مديونية فقط" (لخط اتغطى ولو جزئيًا،
    يعني فيه عملاء تانيين من نفس الخط ظهروا في PDF) و"الخط لم يُرسل"
    (خط تغطيته صفر تمامًا، يعني مفيش أي PDF اتبعت له أصلاً).
    ترجع (result_معدّل, جدول_التغطية). جدول التغطية بيرجع None لو العمود
    غير متاح، وفي الحالة دي result بيرجع زي ما هو من غير تعديل.
    """
    coverage = compute_route_coverage(result)
    if coverage is None:
        return result, None

    result = result.copy()
    coverage_map = coverage.set_index(ROUTE_COL)["نسبة التغطية %"]
    result["نسبة تغطية الخط %"] = result[ROUTE_COL].map(coverage_map)

    is_debt_only = result["الحالة"] == "مديونية فقط"
    route_not_covered = result["نسبة تغطية الخط %"].fillna(0) == 0
    result.loc[is_debt_only & route_not_covered, "الحالة"] = NOT_SENT_STATUS

    return result, coverage


def compute_overdue_ratio(result, followup=None):
    """
    مؤشرات التجاوز، محسوبة على عملاء PDF فقط (عملاء المديونية اللي مش في
    PDF مالهمش دعوة بالحساب):
    - عدد عملاء PDF، وعدد المتجاوزين منهم، ونسبتهم العددية
    - إجمالي مبلغ التجاوز، وإجمالي صافي المديونية، والنسبة بينهم بالقيمة
    ترجع None لو عمود التجاوز غير متاح في ملف المديونية.
    """
    if OVERDUE_COL not in result.columns:
        return None

    if followup is None:
        followup = build_followup_list(result)

    in_pdf = result["رصيد PDF"].notna()
    pdf_customers = int(in_pdf.sum())
    overdue_customers = 0 if followup is None else len(followup)
    ratio = (overdue_customers / pdf_customers * 100) if pdf_customers else 0.0

    total_overdue_amount = float(result.loc[in_pdf, OVERDUE_COL].fillna(0).sum())
    total_net_debt_amount = float(result.loc[in_pdf, "صافي المديونيه"].fillna(0).sum())
    amount_ratio = (
        (total_overdue_amount / total_net_debt_amount * 100)
        if total_net_debt_amount else None  # None يعني القسمة مستحيلة (الإجمالي صفر)، مش 0%
    )

    return {
        "pdf_customers": pdf_customers,
        "overdue_customers": overdue_customers,
        "ratio": ratio,
        "total_overdue_amount": total_overdue_amount,
        "total_net_debt_amount": total_net_debt_amount,
        "amount_ratio": amount_ratio,
    }
