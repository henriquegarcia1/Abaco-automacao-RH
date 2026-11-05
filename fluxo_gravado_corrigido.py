# Fluxo corrigido pelo programa em 04/11/2026 15:29,
# a partir de fluxo_gravado.py. Reproduza com o botao 'Reproduzir
# fluxo' do programa (passar_mouse, clicar e abrir_item sao do programa).
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
    passar_mouse(page, 'Consultas Adicionais por Servidor')
    passar_mouse(page, 'Relatórios Frequências e afastamentos')
    clicar(page, 'link', 'Frequências do Servidor')
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
