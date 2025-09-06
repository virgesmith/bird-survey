from datetime import datetime
from io import BytesIO

# from streamlit_pdf_viewer import pdf_viewer
import streamlit as st
from google import genai  # type: ignore[import-untyped]
from google.genai import types

from model import BtoSpeciesCode, Surveys
from process import BINARY_FORMATS, extract_impl, transform_impl
from spreadsheet import export_to_excel

MODEL = "gemini-2.5-flash"

st.set_page_config(
    page_title="AI Bird Survey Scanner",
    layout="wide",
    page_icon=":bird:",
    menu_items={
        "Report a bug": "https://github.com/virgesmith/bird-survey",
    },
)


def main() -> None:
    st.title(":owl::eagle::duck: AI Bird Survey Scanner")

    st.text_input(
        "Please provide a valid Gemini API key. If you have a google account you can get a free key [here](https://aistudio.google.com/app/apikey)",
        key="api_key",
    )

    st.markdown("#### Step 1: Upload handwritten forms (pdf format) or spreadsheets")
    st.markdown("#### Step 2: AI converts and collates the input files...")
    st.markdown("#### Step 3: Download the resulting spreadsheet")

    with st.expander("More details..."):
        st.markdown("""The survey route is typically split into 10 segments, each represented on the form
                    divided into left and right sections. Observations are written into the appropriate section using
                    species codes, optionally a number and/or a short description (e.g. "flying").
                    Note that descriptions are (deliberately) *not* rendered in the output.
                    Forms typically look like something this:""")
        st.image("./img/raw_data.png")
        st.markdown("""Spreadsheets have no specific format. The AI will only process the first
                    workbook. Results may vary. The species codes that the AI "knows" about are:""")
        code_list = [f"`{code}`: {code.bird_name}" for code in BtoSpeciesCode]
        cols = st.columns(3)
        for i, col in enumerate(cols):
            col.markdown("\n\n".join(code_list[i::3]))

    st.file_uploader("Upload one or more surveys...", type=BINARY_FORMATS, accept_multiple_files=True, key="files")

    file_payloads = {}
    for file in st.session_state.files:
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
            with st.status("Downloading data...", expanded=True) as status:
                for file_name, file_content in file_payloads.items():
                    status.update(label=f"Scanning {file_name}...", state="running")
                    surveys.append(extract_impl(client, MODEL, file_content))
                status.update(label="Complete", state="complete")

            # spreadsheet_content = transform_impl(surveys)
            spreadsheet = export_to_excel(surveys)

            spreadsheet_content = BytesIO()
            spreadsheet.save(spreadsheet_content)

            st.download_button(
                "Download spreadsheet...",
                data=spreadsheet_content.getbuffer(),
                file_name=f"BirdSurvey{datetime.now().isoformat(timespec='seconds')}.xlsx",
            )
    except Exception as e:
        st.error(e)


if __name__ == "__main__":
    main()
