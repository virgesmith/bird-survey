import json
import os
from pathlib import Path

import pandas as pd
import typer  # type: ignore[import-untyped]
from dotenv import load_dotenv
from google import genai  # type: ignore[import-untyped]
from google.genai import types

from model import CLOUD_KEY, RAIN_KEY, VISIBILITY_KEY, WIND_KEY, SurveyData, Surveys

load_dotenv()


def extract(path: Path) -> None:
    files = path.glob("*.pdf")

    prompt = "Convert the information in this pdf file into the requested data structure"
    client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])
    model = os.environ["GEMINI_MODEL"]

    surveys = []
    # for some reason sending all the files at once appears to ignore all but one of them
    for file in files:
        typer.echo(f"Extracting data from {file}")
        with open(file, "rb") as fd:
            file_payload = types.Part.from_bytes(data=fd.read(), mime_type="application/pdf")
        messages = [prompt, file_payload]
        response = client.models.generate_content(
            model=model,
            contents=messages,
            config={
                "response_mime_type": "application/json",
                "response_schema": SurveyData,
            },
        )
        surveys.append(response.parsed)
    output_file = path / "processed_surveys.json"
    with open(output_file, "w") as fd:
        fd.write(Surveys(surveys).model_dump_json(indent=2))
    typer.echo(f"Wrote extracted data to {output_file}")


def transform(path: Path) -> None:
    with open(path / "processed_surveys.json") as fd:
        all_surveys = Surveys(json.load(fd))

    output_file = path / "processed_surveys.xlsx"
    typer.echo(f"Writing transformed data to {output_file}...")

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        for survey_data in all_surveys:
            sheet_name = f"Transect {survey_data.transect_number}"
            typer.echo(f"Creating sheet {sheet_name}")

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

            sighting_data = pd.DataFrame(all_sightings)

            sighting_data.columns = [col.replace("_", " ").capitalize() for col in sighting_data.columns]

            # Create the pandas DataFrame
            sighting_data.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=3,
                startcol=4,
                index=False,
                header=True,
            )

            workbook = writer.book
            header_format = workbook.add_format(
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
            worksheet.set_column(4, 4 + len(sighting_data.columns), 15)
    typer.echo("Done")


def main(path: Path) -> None:
    #extract(path)
    transform(path)


if __name__ == "__main__":
    typer.run(main)
