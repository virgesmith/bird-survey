import json
import os
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import cast

import typer  # type: ignore[import-untyped]
from dotenv import load_dotenv
from google import genai  # type: ignore[import-untyped]
from google.genai import types
from itrx import Itr

from model import Surveys
from spreadsheet import export_to_excel

FORMATS = {"pdf": "application/pdf"}
# image support? png: "image/png"

PROMPT = "Convert the information in this pdf file into the requested data structure. "
"Dates will be in UK format. Ensure the visit_date field is formatted as YYYY-MM-DD"


def get_payload(file: Path) -> types.Part:
    # typer.echo(f"Extracting data from {file}")
    if file.suffix[1:] not in FORMATS:
        raise ValueError(f"{file} is not a valid type")
    with open(file, "rb") as fd:
        file_payload = types.Part.from_bytes(data=fd.read(), mime_type=FORMATS[file.suffix[1:]])
    return file_payload


def extract(path: Path, client: genai.Client, model: str) -> Path:
    files: list[Path] = []
    for ext in FORMATS:  # + TEXT_FORMATS:
        files.extend(path.glob(f"*{ext}"))

    # If you give it a lot of file it seems to miss stuff
    file_contents = Itr(get_payload(file) for file in files).batched(5)
    surveys = Surveys([])

    try:
        # surveys = Surveys([extract_impl(client, model, file_content) for file_content in file_contents])
        for chunk in file_contents:
            print("." * len(chunk))
            surveys.extend(extract_impl(client, model, chunk))
    finally:
        # dump output
        output_file = Path(f"{path}_processed_{datetime.now().isoformat(timespec='seconds')}.json")
        with open(output_file, "w") as fd:
            fd.write(surveys.model_dump_json(indent=2))  # type: ignore
        typer.echo(f"Wrote extracted data to {output_file}")
    return output_file


def extract_impl(client: genai.Client, model: str, file_content: list[types.Part]) -> Surveys:
    # for some reason sending all the files at once appears to ignore all but one of them
    messages = [PROMPT, *file_content]

    response = client.models.generate_content(
        model=model,
        contents=messages,  # type: ignore[arg-type]
        config={
            "response_mime_type": "application/json",
            "response_schema": Surveys,
        },
    )

    return cast(Surveys, response.parsed)


def main(path: Path) -> None:
    load_dotenv()

    client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])
    model = os.environ["GEMINI_MODEL"]

    print(f"Extracting data from {path}")
    extract_data_file = Path(extract(path, client, model))
    print(f"Raw data saved to {extract_data_file}")
    # extract_data_file = Path("data/input_processed_2025-09-07T09:26:54.json")
    with open(extract_data_file) as fd:
        surveys = Surveys(json.load(fd))
    workbook = export_to_excel(surveys)

    buffer = BytesIO()
    workbook.save(buffer)

    spreadsheet = extract_data_file.with_suffix(".xlsx")
    print(f"Output saved to {spreadsheet}")
    with open(spreadsheet, "wb") as fd:
        fd.write(buffer.getvalue())


if __name__ == "__main__":
    typer.run(main)
