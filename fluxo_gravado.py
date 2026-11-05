import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://mgnl.abaco.com.br/eturmalina/servlet/hwlogin")
    page.locator("tr:nth-child(2) > td:nth-child(2)").click()
    page.locator("#vUSULOGIN").click()
    page.locator("#vUSULOGIN").fill("9722")
    page.locator("#vUSUSENHA").click()
    page.locator("#vUSUSENHA").fill("645120")
    page.locator("#vUSUSENHA").press("Enter")
    page.goto("https://mgnl.abaco.com.br/eturmalina/servlet/hwinteracaosistema")
    page.get_by_role("button", name="GESTÃO DE PESSOAS").click()
    page.locator("#ext-gen99").get_by_role("cell", name="Consultas e Relatórios").click()
    page.get_by_role("link", name="Frequências do Servidor").click()
    page.locator("#vCONTRATOMATRICULAINICIAL").click()
    page.locator("#vCONTRATOMATRICULAINICIAL").fill("9722")
    page.locator("#vCONTRATOMATRICULAFINAL").click()
    page.locator("#vCONTRATOMATRICULAFINAL").fill("9723")
    page.locator("#vSERVIDORNOME").click()
    page.locator("#vSERVIDORNOME").press("CapsLock")
    page.locator("#vSERVIDORNOME").fill("G")
    page.locator("#vSERVIDORNOME").press("CapsLock")
    page.locator("#vSERVIDORNOME").fill("Gilson")
    page.get_by_role("button", name="Consultar").click()
    page.close()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
