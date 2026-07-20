import streamlit as st

from collections_core import (
    build_collections_excel,
    extract_collector_label,
    parse_transaction_pdf,
    summarize_by_customer,
)

st.set_page_config(page_title="التحصيل", page_icon="🧾", layout="wide")

st.title("🧾 التحصيل")
st.caption(
    "صفحة منفصلة تمامًا عن أداة المقارنة - بتحلل ملفات PDF من نوع "
    "'سجل المعاملات' أو 'الحساب ديالك' (كشوف تحصيل يومية)، وتفصل اسم "
    "العميل عن رقمه، وتجمع إجمالي المحصّل منه، وإجمالي كل ملف لوحده"
)

files = st.file_uploader(
    "اختر ملف أو أكتر (سجل المعاملات / الحساب ديالك)",
    type=["pdf"],
    accept_multiple_files=True,
)

run = st.button("▶️ حلل الملفات", use_container_width=True, type="primary")

st.divider()

if run:
    if not files:
        st.error("اختار ملف PDF واحد على الأقل")
    else:
        results = []
        warnings = []
        grand_taken = 0.0
        grand_given = 0.0

        progress_text = st.empty()
        progress_bar = st.progress(0)

        for i, f in enumerate(files):
            progress_text.text(f"⏳ بنقرأ {f.name} ({i + 1}/{len(files)})...")
            try:
                label = extract_collector_label(f)
            except Exception:
                label = f.name

            try:
                tx = parse_transaction_pdf(f)
            except Exception as e:
                warnings.append(("error", f"تعذرت قراءة {f.name}: {e}"))
                progress_bar.progress((i + 1) / len(files))
                continue

            if tx.empty:
                warnings.append(("warning", f"{f.name}: مفيش أي معاملات تم التعرف عليها في الملف ده"))
                progress_bar.progress((i + 1) / len(files))
                continue

            summary = summarize_by_customer(tx)
            total_taken = float(tx["تحصيل (أخذت)"].sum())
            total_given = float(tx["معطى (أعطيت)"].sum())
            grand_taken += total_taken
            grand_given += total_given

            results.append({
                "label": label,
                "summary": summary,
                "total_taken": total_taken,
                "total_given": total_given,
            })
            progress_bar.progress((i + 1) / len(files))

        progress_text.empty()
        progress_bar.empty()

        st.session_state["collections_results"] = results
        st.session_state["collections_warnings"] = warnings
        st.session_state["collections_grand_taken"] = grand_taken
        st.session_state["collections_grand_given"] = grand_given
        st.session_state["collections_excel"] = (
            build_collections_excel(results) if results else None
        )

# ---------------------------------------------------------------------------
# عرض النتائج من session_state - يفضل ظاهر حتى لو دوست تحميل أو أي تفاعل تاني
# ---------------------------------------------------------------------------
if "collections_results" in st.session_state:

    for kind, msg in st.session_state.get("collections_warnings", []):
        if kind == "error":
            st.error(f"❌ {msg}")
        else:
            st.warning(f"⚠️ {msg}")

    results = st.session_state["collections_results"]

    for item in results:
        st.subheader(f"📄 {item['label']}")
        st.dataframe(item["summary"], use_container_width=True, hide_index=True)

        msg = f"**💰 إجمالي التحصيل في الملف ده: {item['total_taken']:,.2f}**"
        if item["total_given"]:
            msg += f"  —  إجمالي المُعطى: {item['total_given']:,.2f}"
        st.markdown(msg)
        st.divider()

    if results:
        st.subheader("📊 الإجمالي الكلي لكل الملفات")

        gc1, gc2, gc3 = st.columns(3)
        gc1.metric("عدد الملفات", len(results))
        gc2.metric("إجمالي التحصيل", f"{st.session_state['collections_grand_taken']:,.2f}")
        gc3.metric("إجمالي المُعطى", f"{st.session_state['collections_grand_given']:,.2f}")

        st.download_button(
            "💾 تنزيل تقرير التحصيل الكامل (Excel)",
            data=st.session_state["collections_excel"],
            file_name="تقرير_التحصيل.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="download_collections",
        )
