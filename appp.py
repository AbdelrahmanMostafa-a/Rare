from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

from debt_core import (
    OVERDUE_COL,
    ROUTE_COL,
    NOT_SENT_STATUS,
    build_followup_list,
    compute_overdue_ratio,
    compute_status,
    read_debt_file,
    read_pdf,
    refine_debt_only_by_route,
)
from excel_export import build_excel_report

st.set_page_config(
    page_title="DebtCompare Pro",
    page_icon="📊",
    layout="wide"
)

st.title("📊 DebtCompare Pro")
st.caption("برنامج مقارنة المديونيات")

with st.expander("ℹ️ القواعد الحالية", expanded=False):
    st.markdown("""
    - المقارنة برقم العميل فقط
    - ملف المديونية ممكن يكون **Excel أو TXT أو صورة** - النظام بيدور على
      الأعمدة (رقم العميل / اسم العميل / صافي المديونيه) تلقائي في أي مكان،
      مش لازم تكون في صف أو ترتيب معين
    - استخدام صافي المديونيه
    - دعم PDF واحد أو أكثر
    """)

# ---------------------------------------------------------------------------
# رفع الملفات - في الصفحة الرئيسية (مش في العمود الجانبي) عشان ميقفلش لوحده
# ---------------------------------------------------------------------------
st.subheader("📂 رفع الملفات")

up_col1, up_col2 = st.columns(2)

with up_col1:
    debt_file = st.file_uploader(
        "اختر ملف المديونية",
        type=["xlsx", "xls", "txt", "png", "jpg", "jpeg"],
        help="Excel أو TXT أو صورة - المهم إن الأعمدة الثلاثة "
             "(رقم العميل / اسم العميل / صافي المديونيه) تكون موجودة"
    )
    if debt_file:
        st.success(f"✅ تم رفع: {debt_file.name}")

with up_col2:
    pdf_files = st.file_uploader(
        "اختر ملفات PDF",
        type=["pdf"],
        accept_multiple_files=True
    )
    if pdf_files:
        st.success(f"✅ تم رفع {len(pdf_files)} ملف PDF")

run = st.button("▶️ ابدأ المقارنة", use_container_width=True, type="primary")

st.divider()

# ---------------------------------------------------------------------------
# التشغيل: بيحصل مرة واحدة بس لما تدوس الزرار، والنتيجة بتُحفظ في session_state
# عشان تفضل ظاهرة حتى لو استخدمت البحث أو دوست زرار التحميل بعد كده
# ---------------------------------------------------------------------------
if run:

    if debt_file is None:
        st.error("اختار ملف المديونية")

    elif not pdf_files:
        st.error("اختار ملفات PDF")

    else:
        try:
            with st.spinner("⏳ بنقرأ ملف المديونية..."):
                debt, duplicate_count = read_debt_file(debt_file)

            if debt.empty:
                st.error("لم يتم استخراج أي صفوف بيانات من ملف المديونية")
                st.stop()

            pdf_frames = []
            skipped_files = []

            progress_text = st.empty()
            progress_bar = st.progress(0)

            for i, f in enumerate(pdf_files):
                progress_text.text(f"⏳ بنقرأ {f.name} ({i + 1}/{len(pdf_files)})...")
                try:
                    pdf_frames.append(read_pdf(f))
                except Exception as e:
                    skipped_files.append((f.name, str(e)))
                progress_bar.progress((i + 1) / len(pdf_files))

            progress_text.empty()
            progress_bar.empty()

            pdf_frames = [p for p in pdf_frames if not p.empty]

            if not pdf_frames:
                st.error("مفيش أي بيانات تم استخراجها من ملفات PDF")
                st.stop()

            pdf = pd.concat(pdf_frames, ignore_index=True)
            pdf = pdf.groupby("رقم العميل", as_index=False)["رصيد PDF"].max()

            result = debt.merge(pdf, on="رقم العميل", how="outer")

            result["الفرق"] = (
                result["صافي المديونيه"].fillna(0)
                - result["رصيد PDF"].fillna(0)
            )

            result["الحالة"] = result.apply(compute_status, axis=1)
            result, route_coverage = refine_debt_only_by_route(result)

            # نحفظ كل حاجة محتاجينها للعرض في session_state، وبنبني ملف
            # الإكسل مرة واحدة هنا (مش كل مرة تكتب حرف في البحث) لسرعة أفضل
            st.session_state["result"] = result
            st.session_state["route_coverage"] = route_coverage
            st.session_state["debt_preview"] = debt.head(20)
            st.session_state["debt_count"] = len(debt)
            st.session_state["duplicate_count"] = duplicate_count
            st.session_state["skipped_files"] = skipped_files
            st.session_state["followup"] = build_followup_list(result)

            meta = {
                "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "debt_file_name": debt_file.name,
                "pdf_file_names": [f.name for f in pdf_files],
                "debt_rows": len(debt),
            }
            st.session_state["excel_bytes"] = build_excel_report(result, meta=meta)

        except Exception as e:
            st.session_state.pop("result", None)
            st.error(f"❌ حصل خطأ: {e}")

# ---------------------------------------------------------------------------
# عرض النتائج - يعتمد فقط على session_state، فبيفضل ظاهر مهما عملت بعد كده
# (بحث / فلترة / تحميل الإكسل) من غير ما "يرجع الموقع للبداية"
# ---------------------------------------------------------------------------
if "result" in st.session_state:

    result = st.session_state["result"]

    st.success("✅ التقرير جاهز")

    if st.session_state.get("skipped_files"):
        for name, err in st.session_state["skipped_files"]:
            st.warning(f"⚠️ تعذر قراءة الملف {name}: {err}")

    if st.session_state.get("duplicate_count", 0) > 0:
        st.warning(
            f"⚠️ في {st.session_state['duplicate_count']} رقم عميل مكرر في ملف المديونية - "
            "ده ممكن يظهر صفوف مكررة في النتيجة"
        )

    # ----------------------------------------------------------
    # نسبة التجاوز - من عملاء PDF فقط (أهم مؤشر في الداشبورد)
    # ----------------------------------------------------------
    overdue_ratio = compute_overdue_ratio(result, followup=st.session_state.get("followup"))

    if overdue_ratio is not None:
        st.subheader("🎯 نسبة التجاوز (من عملاء PDF فقط)")

        pct = overdue_ratio["ratio"]
        amount_pct = overdue_ratio["amount_ratio"]

        rc1, rc2, rc3 = st.columns(3)
        rc1.metric("عملاء PDF", overdue_ratio["pdf_customers"])
        rc2.metric("منهم متجاوزين", overdue_ratio["overdue_customers"])
        rc3.metric("نسبة عدد العملاء", f"{pct:.1f}%")

        rc4, rc5, rc6 = st.columns(3)
        rc4.metric("إجمالي صافي المديونية", f"{overdue_ratio['total_net_debt_amount']:,.2f}")
        rc5.metric("إجمالي مبلغ التجاوز", f"{overdue_ratio['total_overdue_amount']:,.2f}")
        rc6.metric("نسبة المبلغ", f"{amount_pct:.1f}%" if amount_pct is not None else "—")

        if amount_pct is None:
            st.warning(
                "⚠️ إجمالي صافي المديونية لعملاء PDF = صفر، فمش قادر أحسب نسبة المبلغ. "
                "تأكد إن عمود 'صافي المديونيه' في ملف المديونية فيه قيم حقيقية"
            )

        severity = pct if amount_pct is None else max(pct, amount_pct)

        msg = (
            f"{overdue_ratio['overdue_customers']} من {overdue_ratio['pdf_customers']} عميل "
            f"({pct:.1f}%) من عملاء PDF عندهم تجاوز"
        )
        if amount_pct is not None:
            msg += f"، وده يمثل {amount_pct:.1f}% من إجمالي صافي مديونية عملاء PDF"

        if severity >= 50:
            st.error(f"🔴 {msg} - نتيجة غير مرضية")
        elif severity >= 20:
            st.warning(f"🟠 {msg}")
        else:
            st.success(f"🟢 {msg}")

        st.divider()

    with st.expander(f"👀 معاينة بيانات المديونية المستخرجة ({st.session_state.get('debt_count', 0)} صف)", expanded=False):
        st.dataframe(st.session_state["debt_preview"], use_container_width=True, hide_index=True)
        st.caption("راجع الصفوف دي وتأكد إن القراءة صحيحة - خصوصًا لو الملف كان صورة أو TXT")

    # ----------------------------------------------------------
    # ملخص / لوحة أرقام
    # ----------------------------------------------------------
    st.subheader("📊 ملخص المقارنة")

    counts = result["الحالة"].value_counts()
    total = len(result)
    matched = int(counts.get("مطابق", 0))
    has_diff = int(counts.get("يوجد فرق", 0))
    pdf_only = int(counts.get("PDF فقط", 0))
    debt_only = int(counts.get("مديونية فقط", 0))
    not_sent = int(counts.get(NOT_SENT_STATUS, 0))
    total_diff_amount = (
        result.loc[result["الحالة"] == "يوجد فرق", "الفرق"]
        .abs()
        .sum()
    )

    show_not_sent = ROUTE_COL in result.columns
    metric_cols = st.columns(6 if show_not_sent else 5)
    metric_cols[0].metric("إجمالي العملاء", total)
    metric_cols[1].metric("✅ مطابق", matched)
    metric_cols[2].metric("⚠️ يوجد فرق", has_diff)
    metric_cols[3].metric("📄 PDF فقط", pdf_only)
    metric_cols[4].metric("📋 مديونية فقط", debt_only)
    if len(metric_cols) == 6:
        metric_cols[5].metric("🚫 الخط لم يُرسل", not_sent)

    if has_diff > 0:
        st.caption(f"💰 إجمالي قيمة الفروقات: {total_diff_amount:,.2f}")

    st.bar_chart(counts)

    # ----------------------------------------------------------
    # تغطية الخطوط (لو عمود "الخط" متاح في ملف المديونية)
    # ----------------------------------------------------------
    route_coverage = st.session_state.get("route_coverage")

    if route_coverage is not None:
        with st.expander(f"🗺️ تغطية الخطوط ({len(route_coverage)} خط)", expanded=False):
            st.caption(
                "نسبة عملاء كل خط اللي ظهروا فعلاً في ملفات PDF المرسلة. الخط "
                "اللي تغطيته صفر معناه معندناش أي PDF ليه أصلاً - فمش بيتحسب "
                "ضمن 'مديونية فقط' الحقيقية"
            )
            st.dataframe(route_coverage, use_container_width=True, hide_index=True)

    # ----------------------------------------------------------
    # عملاء يجب المرور عليهم (لو عمود التجاوز متاح في ملف المديونية)
    # ----------------------------------------------------------
    followup_df = st.session_state.get("followup")

    if followup_df is not None:
        st.subheader("🚨 عملاء يجب المرور عليهم")
        st.caption(
            "العملاء الموجودين في ملفات PDF وعندهم تجاوز في المدة المطلوب "
            "سدادها، مرتبين حسب الرصيد الموجود في PDF من الأكبر للأصغر"
        )
        st.metric("عدد العملاء", len(followup_df))
        if len(followup_df) > 0:
            display_cols = ["الترتيب", "رقم العميل", "اسم العميل", OVERDUE_COL, "رصيد PDF", "الحالة"]
            st.dataframe(
                followup_df[display_cols],
                use_container_width=True,
                hide_index=True,
            )
    else:
        st.info(
            "ℹ️ ملف المديونية ده معندهوش عمود 'اجمالي تجاوز المده المطلوب سداده'، "
            "فمش هينشئ شيت 'عملاء يجب المرور عليهم' في التقرير"
        )

    # ----------------------------------------------------------
    # فلاتر وبحث واستعراض النتائج
    # ----------------------------------------------------------
    st.subheader("🔍 استعراض النتائج")

    fc1, fc2 = st.columns([2, 1])

    with fc1:
        search = st.text_input("بحث برقم العميل أو الاسم", key="search_box")

    with fc2:
        status_options = result["الحالة"].unique().tolist()
        status_filter = st.multiselect(
            "فلترة بالحالة",
            options=status_options,
            default=status_options,
            key="status_filter",
        )

    filtered = result[result["الحالة"].isin(status_filter)]

    if search:
        mask = (
            filtered["رقم العميل"].astype(str).str.contains(search, na=False)
            | filtered["اسم العميل"].astype(str).str.contains(search, na=False)
        )
        filtered = filtered[mask]

    st.caption(f"عدد النتائج المعروضة: {len(filtered)}")
    st.dataframe(filtered, use_container_width=True, hide_index=True)

    # ----------------------------------------------------------
    # تصدير التقرير الكامل Excel (تفاعلي: فلاتر + ألوان حسب الحالة)
    # ----------------------------------------------------------
    excel_bytes = st.session_state.get("excel_bytes")
    if excel_bytes is None:
        excel_bytes = build_excel_report(result)
        st.session_state["excel_bytes"] = excel_bytes

    st.download_button(
        "💾 تنزيل التقرير الكامل (Excel)",
        data=excel_bytes,
        file_name="تقرير_المقارنة.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="download_report",
    )
