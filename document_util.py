import os
import re
import time
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from .config import SAVE_BASE_DIR

def extract_folder_info(driver):
    desktop_title_elem = WebDriverWait(driver, 2).until(
        lambda d: (t := d.find_element(By.CSS_SELECTOR, "div.noselect.desktop-title")).text.strip() != "" and t
    )
    desktop_title = desktop_title_elem.text
    match = re.search(f"公文文號：(\\d+)", desktop_title)
    flow_number = match.group(1) if match else "未知文號"

    try:
        dept_elem = WebDriverWait(driver, 2).until(
            lambda d: (t := d.find_element(By.CSS_SELECTOR, "div[data-speed-id='FullName']")).text.strip() != "" and t
        )
        full_text = dept_elem.text.strip()
        text = full_text.split(" ")[0].strip()
        dept = text[:-1] if text.endswith('函') else text
    except:
        dept = "未知單位"

    issue_date = WebDriverWait(driver, 2).until(
        EC.presence_of_element_located((By.XPATH, '//div[contains(@data-speed-label, "發文日期")]'))
    ).text.strip().replace("中華民國", "")

    issue_number = driver.find_element(By.CSS_SELECTOR, 'div[data-speed-label="發文字號："]').text.strip()

    folder_name = f"{flow_number} {dept} {issue_date} {issue_number}"
    for char in r'\/:*?"<>|':
        folder_name = folder_name.replace(char, '')

    folder_path = os.path.join(SAVE_BASE_DIR, dept + f'\\' + folder_name)  # Remove f"\\" 的話就不會根據各部門分類
    os.makedirs(folder_path, exist_ok=True)
    pdf_filename = f"{flow_number} {dept} {issue_date} {issue_number}.pdf"
    pdf_path = os.path.join(folder_path, pdf_filename)

    return flow_number, folder_path, pdf_path

def extract_expected_attachments(driver):
    try:
        attach_text = driver.find_element(By.CSS_SELECTOR, 'div[data-speed-label^="附件"]').text
        matches = re.findall(r'\(([^)]+)\)', attach_text)
        filenames = []
        for match in matches:
            for f in match.strip().split('、'):
                if '.' in f:
                    filenames.append(f)
        logging.info(f"📎 預期附件：{filenames}")
        return filenames
    except Exception as e:
        logging.error(f"⚠️ 附件名稱擷取失敗: {e}")
        return []

def save_attachment(driver, folder_path, expected_files=None, existed_files=None, max_retries=3):
    for attempt in range(max_retries):
        try:
            WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[data-speed-action="attachment"]'))
            ).click()
            time.sleep(1)

            grid = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.ID, "attachment-grid"))
            )
            filenames = grid.find_elements(By.CSS_SELECTOR, 'td[role="gridcell"]')
            found_filenames = [el.text.strip() for el in filenames if el.text.strip().lower().endswith(('.zip', '.ods', '.xlsx', '.docx', '.pdf', '.csv'))]

            logging.info(f"📄 目前畫面檢測到的檔名：{found_filenames}")

            if existed_files:
                already_exist = all(
                    any(found in exist for exist in existed_files)
                    for found in found_filenames
                )
                if already_exist:
                    logging.info("✅ 所有附件檔案都已存在，跳過下載。")
                    return True

            if expected_files:
                all_found = all(
                    any(expected in detected for detected in found_filenames)
                    for expected in expected_files
                )
                if not all_found:
                    logging.info("⚠️ 有預期檔案尚未出現在畫面上")
                    logging.info(f'expected_files: {expected_files}')  
                    logging.info(f'found_filenames: {found_filenames}')
                    driver.refresh()
                    time.sleep(1)
                    continue

            links = grid.find_elements(By.CSS_SELECTOR, 'i.icon-download')
            if not links:
                return False

            before = set(os.listdir(folder_path))
            for icon in links:
                try:
                    icon.find_element(By.XPATH, './ancestor::a[1]').click()
                    time.sleep(1)
                except:
                    continue

            time.sleep(1)
            after = set(os.listdir(folder_path))
            new_files = after - before

            if expected_files and all(any(f in nf for nf in new_files) for f in expected_files):
                return True
            elif not expected_files:
                return bool(new_files)

        except Exception as e:
            logging.error(f"⚠️ 附件儲存錯誤：{e}")
        driver.refresh()
        time.sleep(1)
    return False