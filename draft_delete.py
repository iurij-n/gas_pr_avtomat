import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False, slow_mo=500)
    context = browser.new_context(storage_state="auth.json")
    page = context.new_page()
    page.set_default_timeout(90000)
    page.goto("https://ej.sudrf.ru/")
    page.get_by_title("Информация о движении поданных обращений в суд").click()
    page.get_by_role("button", name="Найти").click()
    page.pause()
    page.get_by_role("button", name="Удалить").click()
    page.get_by_role("button", name="Удалить").click()
    page.close()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
