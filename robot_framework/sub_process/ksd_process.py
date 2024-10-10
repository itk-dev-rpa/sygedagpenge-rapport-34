"""This module handles interactions with KSDP (Kommunernes Sygedagpengesystem)."""

import os
from csv import DictReader
import time
from dataclasses import dataclass
from datetime import date, datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from itk_dev_shared_components.misc import file_util

from robot_framework import config


@dataclass(init=False)
# pylint: disable-next=too-many-instance-attributes
class Case:
    """A dataclass representing a single case."""
    creation_date: date
    case_number: str
    cpr_number: str
    name: str
    cvr_number: str
    company_name: str
    company_type: str
    first_absence_date: date
    last_absence_date: date
    partial_work_resumption_date: date
    partial_incapacity_date: date
    partial_incapacity_status: str
    absence_reason: str
    absence_reason_note: str
    phone_number: str


def login() -> webdriver.Edge:
    """Login to KSDP using SSO and return the browser object.

    Args:
        orchestrator_connection: The connection to Orchestrator.

    Returns:
        A browser logged in to KSDP.
    """
    edge_options = webdriver.EdgeOptions()
    edge_options.add_argument(f"user-data-dir={os.path.expanduser(r"~\AppData\Local\Microsoft\Edge\User Data\Default")}")
    edge_options.add_argument("profile-directory=Default")
    edge_options.add_experimental_option("prefs", {
        "download.default_directory": os.getcwd(),
        "download.prompt_for_download": False
        # "download.directory_upgrade": True,
        # "safebrowsing.enabled": True
    })
    edge_options.add_argument("--headless")
    edge_options.add_argument("--window-position=-2400,-2400")
    browser = webdriver.Edge(options=edge_options)
    browser.implicitly_wait(2)
    browser.get("https://ksdp.dk/start")

    # Select city if necessary
    try:
        select = Select(browser.find_element(By.ID, "SelectedAuthenticationUrl"))
        select.select_by_visible_text("Aarhus Kommune")
        browser.find_element(By.CSS_SELECTOR, 'input[value=OK]').click()
    except NoSuchElementException:
        pass

    # Wait for site to load
    WebDriverWait(browser, 60).until(EC.element_to_be_clickable((By.ID, "MainShell-logout")))

    return browser


def create_report(browser: webdriver.Chrome, year_from: int, week_from: int, year_to: int, week_to: int):
    """Generate a csv report 34 in KSDP.

    Args:
        browser: A browser logged in to KSDP.
        year_from: The year of the from date.
        week_from: The week of the from date.
        year_to: The year of the to date.
        week_to: The week of the to date.
    """
    browser.find_element(By.CSS_SELECTOR, "span[title=Funktioner]").click()
    browser.find_element(By.CSS_SELECTOR, "button[title=Rapporter]").click()
    browser.find_element(By.CSS_SELECTOR, "span[title='R34 Liste over nyoprettede sager']").click()

    # Set dates
    browser.find_element(By.CSS_SELECTOR, "span[id$=--weekCB]").click()

    from_input = browser.find_element(By.CSS_SELECTOR, "input[id$=--fromDP-inner]")
    from_input.clear()
    from_input.send_keys(f"{year_from}-{week_from}")

    to_input = browser.find_element(By.CSS_SELECTOR, "input[id$=--toDP-inner]")
    to_input.clear()
    to_input.send_keys(f"{year_to}-{week_to}")

    # Download
    browser.find_element(By.CSS_SELECTOR, "button[id$=--oReportsHENTCSVSOM]").click()

    # Wait for download
    file_path = file_util.wait_for_download(os.getcwd(), None, ".csv")

    _close_all_tabs(browser)
    return file_path


def read_csv_file(file_path: str) -> list[Case]:
    """Read the rapport 34 csv file, filter relevant cases and
    return them as a list.

    Args:
        file_path: The path to the csv file.

    Returns:
        A list of dicts describing each case.
    """
    cases = []

    with open(file_path, encoding="UTF-8-sig") as file:
        reader = DictReader(file, delimiter=";")
        for line in reader:
            if line['Sygemeldt-Type'] == "Selvstændig" and "Afsluttet" not in line['Sagsstatus'] and "Lukket" not in line['Sagsstatus']:
                c = Case()
                c.creation_date = _convert_date(line['Opret-dato'], "%Y-%m-%d")
                c.case_number = line['Sagsnummer']
                c.cpr_number = line['CPR-nummer']
                c.name = line['Borger']
                c.cvr_number = line['CVR-nummer']
                c.company_name = line['Virksomhed']
                c.first_absence_date = _convert_date(line['Første fraværsdag'], "%Y-%m-%d")
                c.last_absence_date = _convert_date(line['Sidste fraværsdag'], "%Y-%m-%d")
                c.partial_work_resumption_date = _convert_date(line['Delvis genoptaget arbejde'], "%Y-%m-%d")
                cases.append(c)

    return cases


def get_case_info(browser: webdriver.Chrome, _case: Case) -> None:
    """Open the given case in KSDP and fill out missing information.

    Args:
        browser: A browser logged in to KSDP.
        _case: The case object to enrich.
    """
    # Search and open case
    browser.find_element(By.ID, "__jsview0--TFSearchResultCaseNo").clear()
    browser.find_element(By.ID, "__jsview0--TFSearchResultCaseNo").send_keys(_case.case_number)
    browser.find_element(By.ID, "__button0").click()
    WebDriverWait(browser, 10).until(lambda b: b.find_element(By.ID, "__table0-rows-row0-col6").text == _case.case_number)  # Wait for case number to appear
    browser.find_element(By.ID, "__table0-rows-row0-col0").click()
    _wait_for_loading(browser)

    # Get phone number on first page
    _case.phone_number = browser.find_element(By.CSS_SELECTOR, "input[id$=--TelefonnummerTF]").get_attribute("value")

    # Change page
    browser.find_element(By.CSS_SELECTOR, "a[id$=--navbar-2]").click()
    _wait_for_loading(browser)

    # Get info
    _case.absence_reason = browser.find_element(By.CSS_SELECTOR, "input[id$=--DDBFravaersAarsag-input]").get_attribute("value")
    _case.absence_reason_note = browser.find_element(By.CSS_SELECTOR, "textarea[id$=--TFFravaersAarsagBem]").get_attribute("value")

    delvist_uarbejdsdygtig_dato = browser.find_element(By.CSS_SELECTOR, "input[id$=--DPDelvisUarbejdsdygtigStartdato-col0-row0-input]").get_attribute("value")
    _case.partial_incapacity_date = _convert_date(delvist_uarbejdsdygtig_dato, "%d%m%Y")
    _case.partial_incapacity_status = browser.find_element(By.CSS_SELECTOR, "input[id$=--DPDelvisUarbejdsdygtigAndel-col1-row0-input]").get_attribute("value")

    _close_all_tabs(browser)


def _close_all_tabs(browser: webdriver.Chrome):
    """Close all open tabs in KSDP.
    Note the first tab can't be closed.

    Args:
        browser: A browser logged in to KSDP.
    """
    tab_close_buttons = browser.find_elements(By.CLASS_NAME, "kmdtabclose")

    while len(tab_close_buttons) > 1:
        tab_close_buttons[1].click()
        tab_close_buttons = browser.find_elements(By.CLASS_NAME, "kmdtabclose")


def _convert_date(date_string: str, date_string_format: str) -> date | None:
    """Convert a date string from a given format if possible."""
    if date_string:
        return datetime.strptime(date_string, date_string_format).date()

    return None


def _wait_for_loading(browser: webdriver.Chrome):
    """Wait for KSDP to load.
    Detect the loading attribute on the html body.

    Args:
        browser: A browser logged in to KSDP.
    """
    WebDriverWait(browser, 10).until(lambda b: b.find_element(By.TAG_NAME, "body").get_attribute("aria-busy"))  # Wait for loading
    WebDriverWait(browser, 10).until(lambda b: b.find_element(By.TAG_NAME, "body").get_attribute("aria-busy") is None)  # Wait for loading to disappear
    time.sleep(1)  # A little extra wait for the data to load
