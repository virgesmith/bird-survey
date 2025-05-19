from datetime import datetime

# from streamlit_pdf_viewer import pdf_viewer
import streamlit as st
from google import genai  # type: ignore[import-untyped]
from google.genai import types

from model import Surveys
from process import extract_impl, transform_impl, xlsx_to_csv

MODEL = "gemini-2.5-flash-preview-04-17"

def main() -> None:
    st.title(":owl::bird::eagle::duck: Bird Survey Converter")
    st.markdown("Upload handwritten forms (pdf format), compile to a spreadsheet...")

    st.text_input("Please input a valid Gemini API key", key="api_key")
    st.file_uploader("Upload one of more surveys...", type=["pdf", "xlsx"], accept_multiple_files=True, key="files")

    file_payloads = {}
    for file in st.session_state.files:
        if file.name.endswith("xlsx"):
            file_payloads[file.name] = types.Part.from_text(text=xlsx_to_csv(file.getvalue()))
        else:
            file_payloads[file.name] = types.Part.from_bytes(data=file.getvalue(), mime_type=file.type)
    go = st.button("Process...")

    try:
        if go:
            if not file_payloads:
                st.error("No files selected, upload some surveys using the Browse Files button above")
                return
            if not st.session_state.api_key:
                st.error("Please provide an API key")
                return
            client = genai.Client(api_key=st.session_state.api_key)

            surveys = Surveys([])
            progress_bar = st.progress(0, text="")
            n_files = len(file_payloads)
            for i, (file_name, file_content) in enumerate(file_payloads.items()):
                progress_bar.progress(i / n_files, file_name)
                surveys.append(extract_impl(client, MODEL, file_content))
            progress_bar.progress(1.0, "Complete")

            spreadsheet_content = transform_impl(surveys)

            st.download_button(
                "Download spreadsheet...",
                data=spreadsheet_content,
                file_name=f"BirdSurvey{datetime.now().isoformat(timespec='seconds')}.xlsx",
            )
    except Exception as e:
        st.error(e)

if __name__ == "__main__":
    main()
