import os
import time
import logging
import shutil
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from .driver_util import init_driver, login
from .pdf_handler import retry_save_pdf_gui
from .document_util import extract_folder_info, extract_expected_attachments

TEMP_DOWNLOAD_DIR = os.path.abspath(f"OD_Folders\\TEMP_DOWNLOADS")
os.makedirs(TEMP_DOWNLOAD_DIR, exist_ok=True)

def save_attachment(driver, folder_path, expected_files=None, existed_files=None, max_retries=2):
    for attempt in range(max_retries):
        try:
            WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[data-speed-action="attachment"]'))
            ).click()
            time.sleep(1)

            grid = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.ID, "attachment-grid"))
            )
            filenames = grid.find_elements(By.CSS_SELECTOR, 'td[role="gridcell"]')
            found_filenames = [el.text.strip() for el in filenames if el.text.strip().lower().endswith(('.zip', '.ods', '.xlsx', '.docx', '.pdf'))]

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
                logging.info(f"📭 此文件無可下載附件:{folder_path}")
                print(f"📭 此文件無可下載附件:{folder_path}")
                return False

            before = set(os.listdir(TEMP_DOWNLOAD_DIR))
            for icon in links:
                try:
                    icon.find_element(By.XPATH, './ancestor::a[1]').click()
                    time.sleep(1)
                except:
                    continue

            time.sleep(1)
            after = set(os.listdir(TEMP_DOWNLOAD_DIR))
            new_files = after - before

            for f in new_files:
                src = os.path.join(TEMP_DOWNLOAD_DIR, f)
                dst = os.path.join(folder_path, f)
                shutil.move(src, dst)
                logging.info(f"📁 移動附件：{f} → {folder_path}")

            if expected_files and all(any(f in nf for nf in new_files) for f in expected_files):
                return True
            elif not expected_files:
                return bool(new_files)

        except Exception as e:
            logging.error(f"⚠️ 附件儲存錯誤：{e}")
        driver.refresh()
        time.sleep(2)
    return False

def process_documents(start_idx=0, end_idx=0):
    driver = init_driver(download_path=TEMP_DOWNLOAD_DIR)
    login(driver)

    try:
        WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.ID, "close-kendo-modal"))
        )
        close_btn = driver.find_element(By.ID, "close-kendo-modal")
        close_btn.click()
        print("✅ 已關閉系統訊息視窗")
        logging.info("✅ 已關閉系統訊息視窗")

        # 若有遮罩背景也可額外處理
        time.sleep(1)

    except Exception:
        logging.info("ℹ️ 未出現系統訊息視窗，繼續執行")

    driver.find_element(By.ID, "ToDoCaseHandling").click()
    WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "tr.k-master-row")))
    index = 0
    if start_idx:
        start_idx = start_idx - 1  # 調整為從0開始計數
    while True:
        if end_idx and index >= end_idx:
            break
        if index < start_idx:
            index += 1
            continue

        try:
            rows = driver.find_elements(By.CSS_SELECTOR, "tr.k-master-row")
            if index >= len(rows):
                logging.info("✅ 所有待辦處理完畢")
                break

            logging.info(f'🔄 處理第 {index+1} 筆待辦')

            row = rows[index]
            cells = row.find_elements(By.TAG_NAME, "td")
            if cells[5].text.strip():
                logging.info(f"⚠️ 第 {index+1} 筆已合併，跳過")
                index += 1
                continue

            # if cells[13].text.strip() not in ['內政部警政署', '內政部警政署刑事警察局']:
            #     logging.info(f"⚠️ 第 {index+1} 筆不是內政部公文，跳過")
            #     index += 1
            #     continue

            row.find_element(By.CSS_SELECTOR, "a[data-speed-id='docNumber']").click()
            WebDriverWait(driver, 2).until(lambda d: len(d.window_handles) > 1)
            driver.switch_to.window(driver.window_handles[-1])

            flow_number, folder_path, pdf_path = extract_folder_info(driver)
            expected_files = extract_expected_attachments(driver)
            existed_files = [f for f in set(os.listdir(folder_path)) if f.lower().endswith(('.zip', '.ods', '.xlsx', '.docx', '.pdf'))]

            if os.path.exists(pdf_path) and len(os.listdir(folder_path)) >= 1:
                logging.info(f"✅ 已完成：{pdf_path}，跳過")
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                index += 1
                continue

            if retry_save_pdf_gui(driver, pdf_path, row_index=index):
                logging.info(f"✅ 成功儲存公文：{pdf_path}")
            else:
                logging.info(f"❌ 公文儲存失敗：{pdf_path}，跳過")
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                index += 1
                continue

            driver.close()
            driver.switch_to.window(driver.window_handles[0])

            row = driver.find_elements(By.CSS_SELECTOR, "tr.k-master-row")[index]
            row.find_element(By.CSS_SELECTOR, "a[data-speed-id='docNumber']").click()
            WebDriverWait(driver, 3).until(lambda d: len(d.window_handles) > 1)
            driver.switch_to.window(driver.window_handles[-1])

            if save_attachment(driver, folder_path, expected_files, existed_files):
                logging.info(f"✅ 成功儲存附件：{folder_path}")
            else:
                logging.info(f"❌ 附件儲存失敗：{folder_path}")

            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            index += 1

        except Exception as e:
            logging.error(f"❌ 錯誤：{e}")
            if len(driver.window_handles) > 1:
                driver.close()
            driver.switch_to.window(driver.window_handles[0])
            index += 1

    driver.quit()
    logging.info("✔️完成一輪處理")
