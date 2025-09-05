import json
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, PatternFill

from model import SurveyData, Surveys


def export_survey_to_excel(surveys: list[SurveyData], filename: Path):
    wb = openpyxl.Workbook()
    # TODO multiple worksheets
    ws = wb.active
    ws.title = "draft"

    # Header rows
    ws["B1"] = f"Transect number: {filename.name.split('_')[0]}"
    ws["B2"] = "km square grid reference: TODO"

    # Segment headers
    segment_count = 12
    col = 3
    side = "left"
    for seg in range(segment_count * 2):
        ws.cell(row=2, column=col, value=f"segment {seg // 2 + 1} - {side}")
        ws.cell(row=3, column=col, value="species")
        ws.cell(row=3, column=col + 1, value="count")
        col += 2
        side = "left" if side == "right" else "right"

    # Fill visit rows
    base_row = 4
    for visit_idx, survey in enumerate(surveys):
        row = base_row + visit_idx * 14
        ws.cell(row=row, column=1, value=f"Visit {visit_idx + 1}")
        ws.cell(row=row + 1, column=1, value="Date:")
        ws.cell(row=row + 2, column=1, value="surveyors:")
        ws.cell(row=row + 4, column=1, value="start time:")
        ws.cell(row=row + 5, column=1, value="end time:")
        ws.cell(row=row + 6, column=1, value="start time:")
        ws.cell(row=row + 7, column=1, value="end time:")

        # Fill survey metadata
        ws.cell(row=row + 1, column=2, value=survey.visit_date)
        ws.cell(row=row + 2, column=2, value=survey.observer_name)
        ws.cell(row=row + 4, column=2, value=survey.first_segment_start_time)
        ws.cell(row=row + 5, column=2, value=survey.first_segment_end_time)
        ws.cell(row=row + 6, column=2, value=survey.second_segment_start_time)
        ws.cell(row=row + 7, column=2, value=survey.second_segment_end_time)

        # TODO weather...

        # Fill sightings for each segment
        for segment in survey.segments:
            # TODO add grid ref if available
            seg_idx = segment.number
            col_base = -1 + seg_idx * 4
            # Left
            for sight_idx, sight in enumerate(segment.left):
                ws.cell(row=row + 3 + sight_idx, column=col_base, value=sight.code.bird_name)
                ws.cell(row=row + 3 + sight_idx, column=col_base + 1, value=sight.count or 1)
            # Right
            for sight_idx, sight in enumerate(segment.right):
                ws.cell(row=row + 3 + sight_idx, column=col_base + 2, value=sight.code.bird_name)
                ws.cell(row=row + 3 + sight_idx, column=col_base + 3, value=sight.count or 1)

    # format columns and fill
    # Iterate through all columns in the worksheet, use the column_dimensions property to get the column object,
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 18
    ws.column_dimensions["B"].width = 24

    alternate = False
    for col in ws.iter_cols(min_row=3, max_row=3 + len(surveys) * 14, min_col=3, max_col=2 + segment_count * 4):
        bg = "BBBBBB" if alternate else "DDDDDD"
        for cell in col:
            cell.alignment = Alignment(horizontal="left")
            cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
        alternate = not alternate
    wb.save(filename)


# Example usage:


def main() -> None:
    for file in Path("./data").glob("*.json"):
        with open(file) as fd:
            surveys = Surveys(json.load(fd))
            # for s in surveys:
            export_survey_to_excel(surveys, file.with_suffix(".xlsx"))


if __name__ == "__main__":
    main()
