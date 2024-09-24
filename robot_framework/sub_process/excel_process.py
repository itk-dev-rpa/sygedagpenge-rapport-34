"""This moduel handles Excel files."""

from io import BytesIO

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from robot_framework.sub_process.ksd_process import Case


def write_excel(case_list: tuple[Case]) -> BytesIO:
    """Write the given case list to an Excel sheet.

    Args:
        case_list: The list of cases to write.

    Returns:
        An Excel file as a BytesIO object.
    """
    wb = Workbook()
    sheet: Worksheet = wb.active

    header = [
        "Oprettelsesdato",
        "Sagsnummer",
        "CPR-Nummer",
        "Navn",
        "CVR-Nummer",
        "Virksomhed",
        "Virksomhedsform",
        "Første fraværsdag",
        "Sidste fraværsdag",
        "Delvis genoptaget arbejde dato",
        "Delvist uarbejdsdygtig dato",
        "Delvist uarbejdsdygtig",
        "Fraværsårsag",
        "Fraværsårsag bemærkning",
        "Telefonnummer"
    ]
    sheet.append(header)

    for _case in case_list:
        row = [
            _case.creation_date,
            _case.case_number,
            _case.cpr_number,
            _case.name,
            _case.cvr_number,
            _case.company_name,
            _case.company_type,
            _case.first_absence_date,
            _case.last_absence_date,
            _case.partial_work_resumption_date,
            _case.partial_incapacity_date,
            _case.partial_incapacity_status,
            _case.absence_reason,
            _case.absence_reason_note,
            _case.phone_number
        ]
        sheet.append(row)

    file = BytesIO()
    wb.save(file)
    return file
