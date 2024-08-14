"""This module contains the main process of the robot."""

import os
import uuid
from datetime import datetime

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from itk_dev_shared_components.smtp import smtp_util
from itk_dev_shared_components.smtp.smtp_util import EmailAttachment
from itk_dev_shared_components.misc import cvr_lookup

from robot_framework import config
from robot_framework.sub_process import ksd_process, excel_process


def process(orchestrator_connection: OrchestratorConnection) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")
    cvr_creds = orchestrator_connection.get_credential(config.CVR_CREDS)

    browser = ksd_process.login(orchestrator_connection)

    # Create report
    folder = os.getcwd()
    report_path = os.path.join(folder, f"{uuid.uuid4()}.csv")

    year, week_number, _ = datetime.today().isocalendar()

    ksd_process.create_report(browser, year, week_number, year, week_number,  report_path)
    cases = ksd_process.read_csv_file(report_path)
    os.remove(report_path)

    orchestrator_connection.log_info(f"Searching info on {len(cases)} cases.")

    # Get company type on each case
    for c in cases:
        c.virksomhedsform = cvr_lookup.cvr_lookup(c.cvr_nummer, cvr_creds.username, cvr_creds.password).company_type

    # Filter away cases with the wrong company type
    # TODO Check if correct
    cases = [c for c in cases if c.virksomhedsform == "Enkeltmandsvirksomhed"]

    # Get info from ksd
    for c in cases:
        ksd_process.get_case_info(browser, c)

    excel_file = excel_process.write_excel(cases)
    receivers = orchestrator_connection.process_arguments.split(",")
    smtp_util.send_email(receivers, "itk-rpa@mkb.aarhus.dk", f"Sygedagpenge Rapport 34 - uge {week_number}",
                         f"Her er den berigede rapport 34 for uge {week_number}.\n\nVenlig hilsen\nRobotten",
                         smtp_server=config.SMTP_SERVER, smtp_port=config.SMTP_PORT,
                         attachments=[EmailAttachment(excel_file, "Rapport 34.xlsx")])


if __name__ == '__main__':
    conn_string = os.getenv("OpenOrchestratorConnString")
    crypto_key = os.getenv("OpenOrchestratorKey")
    oc = OrchestratorConnection("Rapport 34 test", conn_string, crypto_key, "")
    process(oc)
