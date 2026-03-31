"""
Streamlit front-end for the RKEI Parser.

Upload one or more .docx RKEI forms, click Process, and download the
resulting Excel workbook.
"""

import os
import tempfile

import streamlit as st

from rkei_parser import process_files

# Page configuration
st.set_page_config(
    page_title="RKEI Form Processor",
    page_icon="📄",
    layout="centered",
)

st.title("📄 RKEI Form Processor")
st.markdown(
    """
    **Welcome!** This tool lets you upload one or more **RKEI `.docx` forms**,
    processes them, and gives you a single **Excel file** containing the extracted data.

    ### How to use
    1. Click **Browse files** below (or drag-and-drop) to upload your `.docx` files.
    2. Press the **▶ Process** button.
    3. Once processing is complete, click **Download Excel** to save the result.
    """
)

st.info(
    "📖 **New here?** Please download and read the user guide below before using this tool."
)


@st.cache_data
def _load_readme() -> str:
    try:
        with open("README.md", "r", encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        return ""


_readme_content = _load_readme()

if _readme_content:
    st.download_button(
        label="📥 Download User Guide",
        data=_readme_content,
        file_name="RKEI_Form_Processor_User_Guide.md",
        mime="text/markdown",
    )

st.divider()

uploaded_files = st.file_uploader(
    "Upload your .docx RKEI forms",
    accept_multiple_files=True,
    type=["docx"],
    help="Select one or more .docx files. All files will be processed together.",
)

if uploaded_files:
    st.info(f"✅ {len(uploaded_files)} file(s) ready for processing.")
else:
    st.warning("⬆ Please upload at least one .docx file to get started.")

process_clicked = st.button("▶ Process", disabled=not uploaded_files)

if process_clicked and uploaded_files:
    with tempfile.TemporaryDirectory() as tmp_dir:
        file_paths = []
        for uf in uploaded_files:
            dest = os.path.join(tmp_dir, uf.name)
            with open(dest, "wb") as f:
                f.write(uf.getbuffer())
            file_paths.append(dest)
            st.info(f"📂 Saved temporary file: {uf.name}")

        try:
            with st.spinner("⏳ Processing files — please wait…"):
                excel_bytes = process_files(file_paths)

            st.success("🎉 Processing complete! Click below to download your Excel file.")
            st.download_button(
                label="⬇ Download Excel",
                data=excel_bytes,
                file_name="SPRE_CodedReturns_REF2029.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as exc:
            st.error(
                "❌ An error occurred while processing the uploaded files. "
                "Please check that you uploaded valid RKEI .docx forms and try again."
            )
            st.exception(exc)