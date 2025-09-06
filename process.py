import json
import os
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import cast

import pandas as pd
import typer  # type: ignore[import-untyped]
from dotenv import load_dotenv
from google import genai  # type: ignore[import-untyped]
from google.genai import types
from itrx import Itr

from model import CLOUD_KEY, RAIN_KEY, VISIBILITY_KEY, WIND_KEY, SurveyData, Surveys
from spreadsheet import export_to_excel

BINARY_FORMATS = ["pdf"]

PROMPT = "Convert the information in this pdf file into the requested data structure. "
"Dates will be in UK format. Ensure the visit_date field is formatted as YYYY-MM-DD"


def get_payload(file: Path) -> types.Part | None:
    typer.echo(f"Extracting data from {file}")
    if file.suffix[1:] in BINARY_FORMATS:
        with open(file, "rb") as fd:
            file_payload = types.Part.from_bytes(data=fd.read(), mime_type=f"application/{file.suffix[1:]}")
        return file_payload


def extract(path: Path, client: genai.Client, model: str) -> Path:
    files: list[Path] = []
    for ext in BINARY_FORMATS: # + TEXT_FORMATS:
        files.extend(path.glob(f"*{ext}"))

    file_contents = Itr(get_payload(file) for file in files)

    #surveys = Surveys([extract_impl(client, model, file_content) for file_content in file_contents])
    surveys = extract_impl(client, model, file_contents)

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

    return response.parsed


def transform(input_file: Path) -> None:
    with open(input_file) as fd:
        surveys = Surveys(json.load(fd))

    output_file = input_file.with_suffix(".xlsx")
    typer.echo(f"Writing transformed data to {output_file}...")

    output_content = transform_impl(surveys)

    with open(output_file, "wb") as fd:
        fd.write(output_content)

    typer.echo("Done")


def _sanitise(value: str) -> str:
    """Sanitise a string to be used as a sheet name in Excel."""
    for bad, good in (("/", "-"), ("\\", "-"), ("?", ""), ("*", ""), (":", ""), ("[", "("), ("]", ")")):
        value = value.replace(bad, good)
    return value


def transform_impl(all_surveys: Surveys) -> bytes:
    output_file = BytesIO()

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        for survey_data in all_surveys:
            sheet_name = f"Transect {survey_data.transect_number} ({_sanitise(survey_data.visit_date)})"

            all_sightings = []

            # Extract common metadata
            common_data = pd.Series(
                {
                    "Observer name": survey_data.observer_name,
                    "Transect number": survey_data.transect_number,
                    "Visit date": survey_data.visit_date,
                    "First segment start time": survey_data.first_segment_start_time,
                    "First segment end time": survey_data.first_segment_end_time,
                    "Second segment start time": survey_data.second_segment_start_time,
                    "Second segment end time": survey_data.second_segment_start_time,
                    # Expand weather codes
                    "Cloud": CLOUD_KEY[survey_data.weather_code.cloud],
                    "Rain": RAIN_KEY[survey_data.weather_code.rain],
                    "Wind": WIND_KEY[survey_data.weather_code.wind],
                    "Visibility": VISIBILITY_KEY[survey_data.weather_code.visibility],
                }
            )
            common_data.to_excel(writer, sheet_name=sheet_name, startrow=3, startcol=1, header=False)

            # Iterate through each segment
            for segment in survey_data.segments:
                segment_number = segment.number
                segment_start_coordinate = segment.start_coordinate
                segment_end_coordinate = segment.end_coordinate

                segment_data = {
                    "segment_number": segment_number,
                    "segment_start_coordinate": segment_start_coordinate,
                    "segment_end_coordinate": segment_end_coordinate,
                }

                # Process sightings on the "left" side
                for sighting in segment.left:
                    sighting_data = {
                        "side": "left",
                        "count": sighting.count or 1,
                        "code": sighting.code,
                        "name": sighting.code.bird_name,
                    }
                    # Combine common, segment, and sighting data for this row
                    all_sightings.append({**segment_data, **sighting_data})

                # Process sightings on the "right" side
                for sighting in segment.right:
                    sighting_data = {
                        "side": "right",
                        "count": sighting.count or 1,
                        "code": sighting.code,
                        "name": sighting.code.bird_name,
                    }
                    # Combine common, segment, and sighting data for this row
                    all_sightings.append({**segment_data, **sighting_data})

            sightings_data = pd.DataFrame(all_sightings)

            sightings_data.columns = pd.Index([col.replace("_", " ").capitalize() for col in sightings_data.columns])

            # Create the pandas DataFrame
            sightings_data.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=3,
                startcol=4,
                index=False,
                header=True,
            )

            workbook = writer.book
            header_format = workbook.add_format(  # type: ignore[union-attr]
                {
                    "bold": True,
                    "text_wrap": True,
                    "valign": "top",
                    "fg_color": "#D7E4BC",  # Light green background
                    "border": 1,
                }
            )

            worksheet = writer.sheets[sheet_name]

            worksheet.write(1, 1, "Survey Details", header_format)
            worksheet.write(1, 4, "Observation Details", header_format)

            worksheet.set_column(1, 2, 25)
            worksheet.set_column(3, 3, 5)
            worksheet.set_column(4, 4 + len(sightings_data.columns), 15)

    return output_file.getvalue()


def main(path: Path) -> None:
    load_dotenv()

    client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])
    model = os.environ["GEMINI_MODEL"]

    extract_data_file = extract(path, client, model)
    # extract_data_file = "data/input_processed_2025-09-06T13:36:17.json"
    with open(extract_data_file) as fd:
        surveys = Surveys(json.load(fd))
    workbook = export_to_excel(surveys)

    buffer = BytesIO()
    workbook.save(buffer)

    with open("test.xlsx", "wb") as fd:
        fd.write(buffer.getvalue())

if __name__ == "__main__":
    typer.run(main)
