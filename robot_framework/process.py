"""This module contains the main process of the robot."""

import os
from datetime import datetime, timedelta
from tkinter import filedialog

from itk_dev_shared_components.misc import cvr_lookup
import dotenv

from robot_framework.sub_process import ksd_process, excel_process


cases = []


def process() -> None:
    """Do the primary process of the robot."""

    global cases
    global save_path

    browser = ksd_process.login()

    year, week_number, _ = (datetime.today() - timedelta(weeks=1)).isocalendar()

    if not cases:
        print(f"Henter rapport for uge {week_number}, {year}")
        report_path = ksd_process.create_report(browser, year, week_number, year, week_number)
        cases = ksd_process.read_csv_file(report_path)
        os.remove(report_path)

        print(f"Antal sager: {len(cases)}")

    # Get company type on each case
    print("Henter virksomhedstyper")
    for c in cases:
        if not c.done:
            c.company_type = cvr_lookup.cvr_lookup(c.cvr_number, os.environ['cvr_username'], os.environ['cvr_password']).company_type

    # Get info from ksd
    for c in cases:
        if not c.done:
            print(f"Henter info på sag {c.case_number}")
            ksd_process.get_case_info(browser, c)

    excel_process.write_excel(cases, save_path)
    os.startfile(save_path)


if __name__ == '__main__':
    save_path = filedialog.asksaveasfilename(initialfile="rapport 34.xlsx", defaultextension="xlsx", filetypes=(("Excel", ".xlsx"),))
    if save_path:
        dotenv.load_dotenv()
        for i in range(3):
            try:
                process()
                break
            except Exception as e:
                if i != 2:
                    print("Hov noget gik galt... Prøver igen.", e)
                else:
                    print("Processen fejlede for mange gange... Kontakt udvikler")
    else:
        print("Ingen sti valgt. Afslutter.")

    input("Robotten er færdig. Luk vinduet eller tryk på en knap for at afslutte.")
