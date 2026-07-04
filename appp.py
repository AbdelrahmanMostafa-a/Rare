from io import BytesIO

import pandas as pd
import streamlit as st

from debt_core import compute_status, read_debt_file, read_pdf

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
# الشريط الجانبي: رفع الملفات وبدء المقارنة
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("📂 الملفات")

    debt_file = st.file_uploader(
        "اختر ملف المديونية",
        type=["xlsx", "xls", "txt", "png", "jpg", "jpeg"],
        help="Excel أو TXT أو صورة (screenshot/صورة ممسوحة) - المهم إن الأعمدة "
             "الثلاثة (رقم العميل / اسم العميل / صافي المديونيه) تكون موجودة"
    )
    if debt_file:
        st.success(f"✅ تم رفع: {debt_file.name}")

    pdf_files = st.file_uploader(
        "اختر ملفات PDF",
        type=["pdf"],
        accept_multiple_files=True
    )
    if pdf_files:
        st.success(f"✅ تم رفع {len(pdf_files)} ملف PDF")

    st.divider()
    run = st.button("▶️ ابدأ المقارنة", use_container_width=True, type="primary")


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

            with st.expander(f"👀 معاينة بيانات المديونية المستخرجة ({len(debt)} صف)", expanded=False):
                st.dataframe(debt.head(20), use_container_width=True, hide_index=True)
                st.caption("راجع الصفوف دي وتأكد إن القراءة صحيحة - خصوصًا لو الملف كان صورة أو TXT")

            if duplicate_count > 0:
                st.warning(
                    f"⚠️ في {duplicate_count} رقم عميل مكرر في ملف المديونية - "
                    "ده ممكن يظهر صفوف مكررة في النتيجة"
                )

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

            if skipped_files:
                for name, err in skipped_files:
                    st.warning(f"⚠️ تعذر قراءة الملف {name}: {err}")

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

            st.success("✅ تم إنشاء التقرير")

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
            total_diff_amount = (
                result.loc[result["الحالة"] == "يوجد فرق", "الفرق"]
                .abs()
                .sum()
            )

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("إجمالي العملاء", total)
            c2.metric("✅ مطابق", matched)
            c3.metric("⚠️ يوجد فرق", has_diff)
            c4.metric("📄 PDF فقط", pdf_only)
            c5.metric("📋 مديونية فقط", debt_only)

            if has_diff > 0:
                st.caption(f"💰 إجمالي قيمة الفروقات: {total_diff_amount:,.2f}")

            st.bar_chart(counts)

            # ----------------------------------------------------------
            # فلاتر وبحث واستعراض النتائج
            # ----------------------------------------------------------
            st.subheader("🔍 استعراض النتائج")

            fc1, fc2 = st.columns([2, 1])

            with fc1:
                search = st.text_input("بحث برقم العميل أو الاسم")

            with fc2:
                status_options = result["الحالة"].unique().tolist()
                status_filter = st.multiselect(
                    "فلترة بالحالة",
                    options=status_options,
                    default=status_options,
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
            # تصدير التقرير الكامل Excel
            # ----------------------------------------------------------
            output = BytesIO()

            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for sheet_status in ["مطابق", "يوجد فرق", "PDF فقط", "مديونية فقط"]:
                    result[result["الحالة"] == sheet_status].to_excel(
                        writer,
                        sheet_name=sheet_status,
                        index=False
                    )

            st.download_button(
                "💾 تنزيل التقرير الكامل (Excel)",
                data=output.getvalue(),
                file_name="تقرير_المقارنة.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"❌ حصل خطأ: {e}")
