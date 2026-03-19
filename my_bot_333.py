import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(storage_state="auth.json")
    page = context.new_page()
    page.goto("https://ej.sudrf.ru/")
    page.locator("#service-appeal").get_by_role("link", name="Подать обращение").click()
    page.get_by_role("link", name="Подать обращение").nth(2).click()
    page.get_by_role("link", name="Заявление о вынесении судебного приказа (дубликата)").click()
    page.get_by_role("checkbox", name="Квитанция об уплате государственной пошлины", exact=True).check()
    page.get_by_role("button", name="Добавить файл").nth(2).click()
    page.get_by_role("button", name="Отменить").click()
    page.get_by_role("checkbox", name="Согласен на получение уведомлений на адрес электронной почты: iurij.").check()
    page.get_by_role("checkbox", name="Согласен на получение судебной корреспонденции на адрес, указанный для направлен").check()
    page.get_by_role("heading", name="Уплата госпошлины").click()
    page.get_by_role("button", name="Сформировать обращение").click()
    page.close()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
