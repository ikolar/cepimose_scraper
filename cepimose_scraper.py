#! /usr/bin/env python
#
# Scrape NIJZ cepimo.se Power BI data portal (https://bit.ly/3rSRfcb)
# using a selenium (firefox automation) script.
#
# Works in headless mode as well.
#
# Prerequisites:
#  - python-selenium python package
#  - geckodriver in path

import time

from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By


def scrape_table(browser, graph_aria):
    """Scrape data for a particular graph (aria-label) by opening its
    data table and walking thought it's cells."""

    # open the table for the graph # by right-clicking it
    # and clicking on the "Show as table" context menu item
    selector = ".visualContainer[aria-label='{}']".format(graph_aria)
    graph = browser.find_element_by_css_selector(selector)
    actionChains = ActionChains(browser)
    actionChains.move_to_element(graph).context_click(graph).perform()
    wait_for(browser, "h6.itemLabel")
    try:
        time.sleep(0.5)
        browser.find_element_by_css_selector("h6.itemLabel").click()
    except ElementNotInteractableException:
        # wait a bit longer
        time.sleep(2)
        browser.find_element_by_css_selector("h6.itemLabel").click()

    # wait for the table to load
    wait_for(browser, ".detailVisual .bodyCells")

    # grab title
    if graph_aria != " Shape":
        title = browser.execute_script(
            """return document.querySelector("div.title.trimmedTextWithEllipsis").textContent;"""
        )
    else:
        title = "Dobave cepiv"

    # we'll be putting the table data here
    rows = []

    # grab column names
    corners = browser.execute_script(
        """return Array.from(document.querySelectorAll("detail-visual-modern .corner > div > .pivotTableCellWrap"), cell => cell.textContent.trim());"""
    )
    columns = browser.execute_script(
        """return Array.from(document.querySelectorAll("detail-visual-modern .columnHeaders > div > div"), cell => cell.textContent.trim());"""
    )
    if len(corners) > 1:
        # columns names are grouped by doses
        # renamed them to ["Odmerek 1 - ...", "Odmerek 1 - ...", "Odmerek 2 - ...", ...]
        row = [
            corners[-1],
        ]
        row.append("Odmerek 1 - " + columns[1])
        row.append("Odmerek 1 - " + columns[2])
        row.append("Odmerek 2 - " + columns[4])
        row.append("Odmerek 2 - " + columns[5])
    else:
        row = [
            corners[0],
        ] + columns
    rows.append(row)

    # move to the first cell in the first row
    time.sleep(2)
    body = browser.find_element_by_tag_name("body")
    if graph_aria != " Shape":
        for i in range(3):
            body.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.05)
        if len(corners) > 1:
            # we need one more arrow down because of the graph legend
            body.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.05)
        body.send_keys(Keys.ARROW_LEFT)
    else:
        # the table with vaccine shipments has another table above it
        # so we need a different set of keystrokes to reach it
        for i in range(2):
            body.send_keys(Keys.ARROW_LEFT)
            time.sleep(0.05)
    time.sleep(2)

    # now, move right cell-by-cell, grabbing the value of the current cell
    # while a bit slow, this seems to be the best way to do it since
    # the data table is dynamically loaded as you scroll down
    while True:
        row = []

        # load in first cell in row
        date = browser.execute_script(
            """return document.querySelector(".hasFocus").textContent.trim();"""
        )
        if date in corners:
            # break if we're back at the start of the table
            break
        row.append(date)

        num_columns = len(rows[0]) - 1
        for i in range(num_columns):
            body.send_keys(Keys.ARROW_RIGHT)
            value = browser.execute_script(
                """return document.querySelector(".hasFocus").textContent.replace(",", "").trim();"""
            )
            if not value:
                value = 0
            elif "." in value:
                value = float(value)
            else:
                value = int(value)

            row.append(value)
        rows.append(row)
        body.send_keys(Keys.ARROW_RIGHT)

    print(title, rows)

    # return to main page
    browser.execute_script("""document.querySelector(".menuItem").click();""")
    time.sleep(3)


def wait_for(browser, css_selector, timeout=15):
    """Wait for a specific page element to show up"""
    element_present = EC.presence_of_element_located((By.CSS_SELECTOR, css_selector))
    WebDriverWait(browser, timeout).until(element_present)


def next_page(browser):
    """Click the "Naslednja Stran" arrow at the bottom of the page."""
    browser.execute_script(
        """document.querySelector("i[title='Next Page']").click();"""
    )


def init(headless=True):
    """Prepare selenium firefox instance."""
    opts = Options()
    if headless:
        opts.set_headless()
        assert opts.headless
    browser = Firefox(options=opts)

    url = "https://app.powerbi.com/view?r=eyJrIjoiZTg2ODI4MGYtMTMyMi00YmUyLWExOWEtZTlmYzIxMTI2MDlmIiwidCI6ImFkMjQ1ZGFlLTQ0YTAtNGQ5NC04OTY3LTVjNjk5MGFmYTQ2MyIsImMiOjl9&pageName=ReportSectionf7478503942700dada61"
    browser.get(url)

    wait_for(
        browser,
        ".visualContainer[aria-label='Skupno število cepljenih oseb Line chart']",
    )

    return browser


if __name__ == "__main__":
    browser = init(headless=False)

    scrape_table(
        browser,
        "Delež cepljenih oseb po starostnih razredih Line and clustered column chart",
    )
    scrape_table(browser, "Skupno število cepljenih oseb Line chart")
    scrape_table(
        browser, "Delež cepljenih oseb po statističnih regijah Clustered bar chart"
    )

    next_page(browser)
    wait_for(browser, ".visualContainer[aria-label='Pogled po cepivu Slicer']")

    scrape_table(browser, " Shape")
    scrape_table(browser, "Skupno število odmerkov cepiva po datumu dobave Line chart")
    scrape_table(browser, "Število dobljenih in porabljenih odmerkov cepiva Area chart")
