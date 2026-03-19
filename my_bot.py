# import re
from playwright.sync_api import Playwright, sync_playwright, expect
import pygetwindow as gw
import pyautogui
import time

try:
    from creditor_data import CREDITOR_DATA
except ImportError:
    CREDITOR_DATA = dict()


DRAFT_TARGET_TYPE = "Заявление о вынесении судебного приказа (дубликата)"
DRAFT_TARGET_STATUS = "В процессе создания (черновик)"
GET_ELEMENT_TIMEOUT = 15000


def close_dialog_if_exists():
    for _ in range(20):
        windows = gw.getWindowsWithTitle('Открытие')
        if windows:
            try:
                windows[0].activate()
            except Exception:
                pass
            else:
                pyautogui.press('esc')
                break
        time.sleep(0.5)


def get_element(page, identificator):
    el = page.locator(identificator)
    try:
        el.wait_for(state="visible", timeout=GET_ELEMENT_TIMEOUT)
    except:
        return None
    else:
        return el


def delete_first_matching_draft(page):

    try:
        page.goto("https://ej.sudrf.ru/")
    except:
        page.reload()

    page.get_by_title("Информация о движении поданных обращений в суд").click()
    page.get_by_role("button", name="Найти").click()

    try:
        table_container = page.locator(".table-responsive")
        table_container.wait_for(state="visible", timeout=15000)
    except:
        return

    rows = page.locator(".table-history tbody tr")
    rows_count = rows.count()

    for i in range(rows_count):
        row = rows.nth(i)
        row_type = row.locator("td").nth(3).inner_text().strip()
        row_status = row.locator("td").nth(4).inner_text().strip()

        if DRAFT_TARGET_TYPE in row_type and DRAFT_TARGET_STATUS in row_status:
            btn_delete = row.locator("td").nth(4).get_by_role("button", name="Удалить")
            if btn_delete.is_visible():
                btn_delete.click()
                confirm_btn = page.locator(".modal-content:has-text('Подтвердите удаление')").get_by_role("button", name="Удалить")
                if confirm_btn.is_visible(timeout=2000):
                    time.sleep(0.5)
                    confirm_btn.click()
                    time.sleep(0.5)
                return
        return


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False, slow_mo=1000)
    context = browser.new_context(storage_state="auth.json")
    page = context.new_page()
    page.set_default_timeout(90000)
    delete_first_matching_draft(page)
    try:
        page.goto("https://ej.sudrf.ru/")
    except:
        page.reload() # Если не вышло — просто обновляем страницу
    page.locator("#service-appeal").get_by_role("link", name="Подать обращение").click()
    page.get_by_role("link", name="Подать обращение").nth(2).click()
    page.get_by_role("link", name="Заявление о вынесении судебного приказа (дубликата)").click()
    # page.pause()
    # page.get_by_role("button", name="Я являюсь представителем").click()

    time.sleep(0.5)
    page.locator('button[name="Method"][value="2"]').click() # Кнопка "Я являюсь представителем"
    # page.locator("#Address_CourtNotices_Index").click()
    # page.locator("label").filter(has_text="Адрес для направления судебных повесток и иных судебных извещений:").click()
    # page.locator("#Address_CourtNotices_Index").click()
    page.locator("#Address_CourtNotices_Index").fill(CREDITOR_DATA.get('NOTIFY_POSTAL_CODE', ''))
    # page.get_by_role("textbox", name="Адрес для направления судебных повесток и иных судебных извещений:").click()
    # page.get_by_role("textbox", name="Адрес для направления судебных повесток и иных судебных извещений:").click()
    # page.get_by_role("textbox", name="Адрес для направления судебных повесток и иных судебных извещений:").press("ControlOrMeta+a")
    page.get_by_role("textbox", name="Адрес для направления судебных повесток и иных судебных извещений:").fill(CREDITOR_DATA.get('NOTIFY_ADDRESS', ''))
    # page.pause()
    page.get_by_role("button", name="Добавить файл").first.click()
    page.get_by_role("button", name="Выбрать файл").click()
    # page.get_by_role("button", name="Выбрать файл").set_input_files("cardxxxx.pdf")
    # page.get_by_role("button", name="Выбрать файл").set_input_files(r"C:\Users\triko\Downloads\Telegram Desktop\Desktop\Файлы_досье\Досье_8092417_Абеляшев Александр Анатольевич\cardxxxx.pdf")
    # 1. Запускаем ожидание события выбора файла
    with page.expect_file_chooser() as fc_info:
        # 2. Кликаем по вашей кнопке, которая "не является инпутом"
        # page.get_by_role("button", name="Выбрать файл").click()
        page.get_by_role("button", name="Выбрать файл").dispatch_event("click")

    # 3. Объект chooser сам найдет нужный скрытый инпут и вставит туда файл
    file_chooser = fc_info.value
    file_chooser.set_files(CREDITOR_DATA.get('AUTHORITY_CONFIRMATION_DOC', r''))
    page.get_by_role("button", name="Добавить").click()
    close_dialog_if_exists()
    page.pause()
    page.get_by_role("button", name="Добавить заявителя").click()
    page.get_by_role("button", name="Юридическое лицо").click()
    # page.get_by_role("textbox", name="Наименование:").click()
    page.get_by_role("textbox", name="Наименование:").fill(CREDITOR_DATA.get('CREDITOR_NAME', ''))
    page.get_by_label("Процессуальный статус заявителя:").select_option("50710012")
    # page.get_by_role("textbox", name="ИНН:").click()
    page.get_by_role("textbox", name="ИНН:").fill(CREDITOR_DATA.get('CREDITOR_INN', ''))
    # page.get_by_role("textbox", name="ИНН:").click()
    # page.get_by_role("textbox", name="ОГРН:").click()
    page.get_by_role("textbox", name="ОГРН:").fill(CREDITOR_DATA.get('CREDITOR_OGRN', ''))
    # page.get_by_role("textbox", name="КПП:").click()
    page.get_by_role("textbox", name="КПП:").fill(CREDITOR_DATA.get('CREDITOR_KPP', ''))
    # page.locator("#Address_Legal_Index").click()
    page.locator("#Address_Legal_Index").fill(CREDITOR_DATA.get('LEGAL_POSTAL_CODE', ''))
    # page.locator("#Address_Legal_Index").press("Tab")
    page.get_by_role("textbox", name="Юридический адрес:").fill(CREDITOR_DATA.get('LEGAL_ADDRESS', ''))
    page.get_by_role("checkbox", name="Адрес фактического нахождения организации совпадает с юридическим адресом органи").check()
    # page.get_by_role("dialog", name="Данные заявителя").locator("#Email").click()
    page.get_by_role("dialog", name="Данные заявителя").locator("#Email").fill(CREDITOR_DATA.get('CREDITOR_EMAIL', ''))
    # page.get_by_role("dialog", name="Данные заявителя").locator("#Phone").click()
    page.get_by_role("dialog", name="Данные заявителя").locator("#Phone").fill(CREDITOR_DATA.get('CREDITOR_PHONE_NUMBER', ''))
    page.get_by_role("button", name="Сохранить").click()
    page.pause()
    page.get_by_role("button", name="Добавить участника").click()
    page.get_by_role("button", name="Физическое лицо").click()
    page.get_by_role("textbox", name="Фамилия").click()
    page.get_by_role("textbox", name="Фамилия").fill("Иванов")
    page.get_by_role("textbox", name="Фамилия").press("Tab")
    page.get_by_role("textbox", name="Имя").fill("Иван")
    page.get_by_role("textbox", name="Имя").press("Tab")
    page.get_by_role("textbox", name="Отчество").fill("Иванович")
    page.get_by_role("textbox", name="Отчество").press("Tab")
    page.get_by_role("textbox", name="Введите дату рождения представляемого лица").fill("01.01.2001")
    page.get_by_role("button", name="Мужской").click()
    page.get_by_role("textbox", name="Введите адрес места рождения").click()
    page.get_by_role("textbox", name="Введите адрес места рождения").fill("32154, Москва 12")
    page.get_by_role("dialog", name="Данные участника процесса").locator("#Address_Permanent_Index").click()
    page.get_by_role("dialog", name="Данные участника процесса").locator("#Address_Permanent_Index").fill("321654.")
    page.get_by_role("dialog", name="Данные участника процесса").locator("#Address_Permanent_Address").click()
    page.get_by_role("dialog", name="Данные участника процесса").locator("#Address_Permanent_Address").fill("Москва 12")
    page.get_by_role("dialog", name="Данные участника процесса").locator("#Address_Actual_Index").click()
    page.get_by_role("dialog", name="Данные участника процесса").locator("#Address_Actual_Index").fill("321654")
    page.get_by_role("dialog", name="Данные участника процесса").locator("#Address_Actual_Address").click()
    page.get_by_role("dialog", name="Данные участника процесса").locator("#Address_Actual_Address").fill("Москва 12")
    page.get_by_role("dialog", name="Данные участника процесса").locator("#Type").select_option("1")
    page.get_by_role("dialog", name="Данные участника процесса").locator("#Series").click()
    page.get_by_role("dialog", name="Данные участника процесса").locator("#Series").fill("32 131")
    page.get_by_role("dialog", name="Данные участника процесса").locator("#Number").click()
    page.get_by_role("dialog", name="Данные участника процесса").locator("#Number").fill("7417414")
    page.get_by_role("textbox", name="Дата выдачи").click()
    page.get_by_role("textbox", name="Дата выдачи").fill("20.10.2010")
    page.get_by_role("dialog", name="Данные участника процесса").locator("#IssuedBy").click()
    page.get_by_role("dialog", name="Данные участника процесса").locator("#IssuedBy").fill("МВД Курск")
    page.get_by_role("dialog", name="Данные участника процесса").locator("#IssuerId").click()
    page.get_by_role("dialog", name="Данные участника процесса").locator("#IssuerId").fill("741-741_")
    page.get_by_role("checkbox", name="Полные данные участника процесса неизвестны").check()
    page.get_by_role("button", name="Сохранить").click()
    page.get_by_role("button", name="Выбрать суд").click()
    page.get_by_label("Регион:").select_option("35")
    page.get_by_label("Судебный орган:").select_option("35MS0046")
    page.get_by_role("button", name="Сохранить").click()
    # page.get_by_role("link", name="Удалить файл").nth(2).click()
    page.get_by_role("button", name="Добавить файл").nth(1).click()
    page.get_by_role("button", name="Выбрать файл").click()
    page.get_by_role("button", name="Выбрать файл").set_input_files("Фото.pdf")
    page.get_by_role("button", name="Добавить").click()
    # page.get_by_role("link", name="Удалить файл").nth(3).click()
    # page.get_by_role("link", name="Удалить файл").nth(3).click()
    # page.get_by_role("link", name="Удалить файл").nth(3).click()
    page.get_by_role("button", name="Добавить файл").nth(1).click()
    page.get_by_role("button", name="Выбрать файл").click()
    page.get_by_role("button", name="Выбрать файл").set_input_files("ПЛАТЕЖНОЕ ПОРУЧЕНИЕ.pdf")
    page.get_by_role("button", name="Добавить").click()
    page.get_by_role("button", name="Добавить файл").nth(1).click()
    page.get_by_role("button", name="Выбрать файл").click()
    page.get_by_role("button", name="Выбрать файл").set_input_files("Уведомление о состоявшейся уступке права требования.pdf")
    page.get_by_role("button", name="Добавить").click()
    page.get_by_role("button", name="Добавить файл").nth(1).click()
    page.get_by_role("button", name="Выбрать файл").click()
    page.get_by_role("button", name="Выбрать файл").set_input_files("ПЛАТЕЖНОЕ ПОРУЧЕНИЕ.pdf")
    page.get_by_role("button", name="Добавить").click()
    page.get_by_role("checkbox", name="Квитанция об уплате государственной пошлины", exact=True).uncheck()
    page.get_by_role("checkbox", name="Квитанция об уплате государственной пошлины", exact=True).check()
    page.locator(".col-sm-6 > div:nth-child(2) > .pull-right").click()
    page.get_by_role("button", name="Добавить файл").nth(2).click()
    page.get_by_role("button", name="Выбрать файл").click()
    page.get_by_role("button", name="Выбрать файл").set_input_files("Паспорт_регистрация.pdf")
    page.get_by_role("button", name="Добавить").click()
    # page.get_by_role("checkbox", name="Согласен на получение уведомлений на адрес электронной почты: iurij.").uncheck()
    page.get_by_role("checkbox", name="Согласен на получение уведомлений на адрес электронной почты: iurij.").check()
    page.get_by_role("checkbox", name="Согласен на получение судебной корреспонденции на адрес, указанный для направлен").uncheck()
    # page.get_by_role("checkbox", name="Согласен на получение судебной корреспонденции на адрес, указанный для направлен").check()
    page.get_by_role("button", name="Сформировать обращение").click()
    page.close()

    # ---------------------
    context.close()
    browser.close()


if __name__ == '__main__':
    with sync_playwright() as playwright:
        run(playwright)
