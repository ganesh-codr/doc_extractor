import tempfile
from pathlib import Path

import streamlit as st

from doc_extractor.extractor import (
    build_schedule_rows,
    export_excel,
    export_pdf,
    extract_text,
    find_poppler_path,
    paginate_rows,
)

st.set_page_config(page_title="Engineering Drawing Extractor", layout="wide")
st.title("Engineering Drawing Extractor")
st.caption("Upload a PDF or text drawing, enter a prompt, and extract structured output into a table.")

with st.expander("Need help with scanned PDFs?", expanded=False):
    st.markdown(
        """
        If your drawing is a scanned PDF, the extractor needs Poppler installed on your machine.
        The app looks for Poppler in this order: the `POPPLER_PATH` environment variable,
        your system PATH, then common install folders like `C:\\poppler-xx\\Library\\bin`.
        """
    )

    if st.button("Test Poppler Path", help="Check whether Poppler tools are visible to the app."):
        try:
            poppler_path = find_poppler_path()
            if poppler_path:
                st.success(f"Poppler found at: {poppler_path}")
            else:
                st.success("Poppler is on PATH. The app can find pdftoppm.")
        except RuntimeError as exc:
            st.error(str(exc))

    st.link_button(
        "Install Poppler for Windows",
        "https://github.com/oschwartz10612/poppler-windows/releases",
        help="Open the Poppler Windows release page to download and install Poppler.",
    )

uploaded_file = st.file_uploader("Upload drawing", type=["pdf", "txt", "png", "jpg", "jpeg"])
prompt = st.text_area(
    "Prompt",
    value="Prepare a cable schedule from the drawing with the columns Sl.no, Cable tag, From Equipment, Location, To Equipment, Location, Cable Size, Cable Type, Remarks.",
    height=120,
)

if st.button("Extract", type="primary"):
    if uploaded_file is None:
        st.warning("Please upload a drawing file before extracting.")
    elif not prompt.strip():
        st.warning("Please provide a prompt describing the extraction you want.")
    else:
        temp_dir = Path(tempfile.gettempdir()) / "doc_extractor_ui"
        temp_dir.mkdir(exist_ok=True)
        temp_path = temp_dir / uploaded_file.name
        temp_path.write_bytes(uploaded_file.getvalue())

        source_type = "txt" if temp_path.suffix.lower() in {".txt", ".md"} else "pdf"
        text = extract_text(temp_path, source_type=source_type)
        rows = build_schedule_rows(text, prompt)

        st.session_state["rows"] = rows
        st.session_state["source_name"] = uploaded_file.name

if "rows" in st.session_state and st.session_state["rows"]:
    rows = st.session_state["rows"]
    total_rows = len(rows)

    st.subheader("Extraction result")
    st.write(f"{total_rows} record(s) found from {st.session_state.get('source_name', 'uploaded drawing')}")

    if total_rows > 25:
        page_count = (total_rows + 24) // 25
        page_number = st.number_input("Page", min_value=1, max_value=page_count, value=1, step=1)
        page_result = paginate_rows(rows, page_number=page_number, page_size=25)
        st.info("Showing 25 records per page.")
    else:
        page_number = 1
        page_result = paginate_rows(rows, page_number=page_number, page_size=25)

    st.dataframe(page_result["rows"], use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        excel_path = Path(tempfile.gettempdir()) / "doc_extractor_ui" / "download_result.xlsx"
        export_excel(rows, excel_path)
        with open(excel_path, "rb") as handle:
            st.download_button(
                "Download Excel",
                handle.read(),
                file_name="engineering_extraction.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    with col2:
        pdf_path = Path(tempfile.gettempdir()) / "doc_extractor_ui" / "download_result.pdf"
        export_pdf(rows, pdf_path)
        with open(pdf_path, "rb") as handle:
            st.download_button(
                "Download PDF",
                handle.read(),
                file_name="engineering_extraction.pdf",
                mime="application/pdf",
            )
