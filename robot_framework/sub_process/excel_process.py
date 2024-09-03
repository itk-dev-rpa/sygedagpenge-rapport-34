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
            _case.oprettelsesdato,
            _case.sagsnummer,
            _case.cpr_nummer,
            _case.navn,
            _case.cvr_nummer,
            _case.virksomhed,
            _case.virksomhedsform,
            _case.første_fraværsdag,
            _case.sidste_fraværsdag,
            _case.delvis_genoptaget_arbejde_dato,
            _case.delvist_uarbejdsdygtig_dato,
            _case.delvist_uarbejdsdygtig,
            _case.fraværsårsag,
            _case.fraværsårsag_bemærkning,
            _case.telefonnummer
        ]
        sheet.append(row)

    file = BytesIO()
    wb.save(file)
    return file
