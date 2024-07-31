"""This module contains the main process of the robot."""

import os
import uuid
from datetime import datetime

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from itk_dev_shared_components.smtp import smtp_util
from itk_dev_shared_components.smtp.smtp_util import EmailAttachment

from robot_framework import config
from robot_framework.sub_process import ksd_process, excel_process


def process(orchestrator_connection: OrchestratorConnection) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")

    browser = ksd_process.login(orchestrator_connection)

    # Create report
    folder = os.getcwd()
    report_path = os.path.join(folder, f"{uuid.uuid4}.csv")

    year, week_number, _ = datetime.today().isocalendar()

    ksd_process.create_report(browser, year, week_number, year, week_number,  report_path)
    cases = ksd_process.read_csv_file(report_path)
    os.remove(report_path)

    orchestrator_connection.log_info(f"Searching info on {len(cases)} cases.")

    for c in cases:
        # TODO cvr lookup
        ksd_process.get_case_info(browser, c)

    excel_file = excel_process.write_excel(cases)
    receivers = orchestrator_connection.process_arguments.split(",")
    smtp_util.send_email(receivers, "itk-rpa@mkb.aarhus.dk", f"Sygedagpenge Rapport 34 - uge {week_number}",
                         f"Her er den berigede rapport 34 for uge {week_number}.\n\nVenlig hilsen\nRobotten",
                         smtp_server=config.SMTP_SERVER, smtp_port=config.SMTP_PORT,
                         attachments=[EmailAttachment(excel_file, "Rapport 34.xlsx")])
