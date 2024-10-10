"""This module contains the main process of the robot."""

import os
import uuid
from datetime import datetime, timedelta
from tkinter import filedialog

from itk_dev_shared_components.misc import cvr_lookup
import dotenv

from robot_framework.sub_process import ksd_process, excel_process


def process() -> None:
    """Do the primary process of the robot."""

    save_path = filedialog.asksaveasfilename(initialfile="rapport 34.xlsx", defaultextension="xlsx", filetypes=(("Excel", ".xlsx"),))
    if not save_path:
        return

    browser = ksd_process.login()

    year, week_number, _ = (datetime.today() - timedelta(weeks=1)).isocalendar()

    report_path = ksd_process.create_report(browser, year, week_number, year, week_number)
    cases = ksd_process.read_csv_file(report_path)
    os.remove(report_path)

    # Get company type on each case
    for c in cases:
        c.company_type = cvr_lookup.cvr_lookup(c.cvr_number, os.environ['cvr_username'], os.environ['cvr_password']).company_type

    # Get info from ksd
    for c in cases:
        ksd_process.get_case_info(browser, c)

    excel_process.write_excel(cases, save_path)


if __name__ == '__main__':
    dotenv.load_dotenv()
    process()
