import os
import time
import pyautogui
import pyperclip
import logging
import pygetwindow as gw
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# def save_pdf_gui(pdf_path):
#     pyautogui.FAILSAFE = True
#     time.sleep(1)

#     pyautogui.click(x=1490, y=280)
#     time.sleep(0.5)
#     pyautogui.click(x=1350, y=350)
#     time.sleep(0.5)

#     if not pyautogui.locateCenterOnScreen("pdf_selected.png", confidence=0.6):
#         logging.info("❌ 驗證失敗")
#         pyautogui.screenshot("verify_pdf_failed.png")
#         return False

#     pyautogui.click(x=1350, y=875)

#     time.sleep(1.5)
#     for _ in range(10):
#         if pyautogui.locateOnScreen('win_save.png', confidence=0.4):
#             break
#         time.sleep(1)
#     else:
#         pyautogui.click(x=805, y=735)

#     logging.debug(f"💾 準備貼上的路徑: {pdf_path}")

#     time.sleep(1)
#     pyautogui.hotkey('ctrl', 'a')
#     time.sleep(0.3)
#     pyperclip.copy(pdf_path)
#     pyautogui.hotkey('ctrl', 'v')
#     time.sleep(1)
#     pyautogui.press('enter')
#     logging.info("📂 儲存命令已送出")

#     logging.debug(f"📋 剪貼簿實際內容: {pyperclip.paste()}")

#     for _ in range(20):
#         if os.path.exists(pdf_path):
#             return True
#         time.sleep(0.5)
#     return False


def wait_and_click(image_path, confidence=0.9, timeout=10):
    """等待並點擊畫面上的圖像"""
    start = time.time()
    while time.time() - start < timeout:
        location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
        if location:
            logging.info(f"🖱️ 找到並點擊圖像：{image_path}")
            pyautogui.click(location)
            return True
        time.sleep(0.5)
    pyautogui.screenshot(f"error_{os.path.basename(image_path)}_{int(time.time())}.png")
    raise Exception(f"❌ 找不到圖像：{image_path}")

def wait_for_image(image_path, confidence=0.9, timeout=10):
    """只等圖片出現，不點擊"""
    start = time.time()
    while time.time() - start < timeout:
        location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
        if location:
            print(f"✅ 找到圖片: {image_path}")
            return location
        time.sleep(0.5)
    return None

def focus_latest_chrome_window():
    time.sleep(0.5)
    windows = [w for w in gw.getWindowsWithTitle(" - Google Chrome")]
    if not windows:
        logging.warning("⚠️ 找不到 Chrome 視窗")
        return
    win = sorted(windows, key=lambda w: w._hWnd)[-1]
    win.activate()
    logging.info(f"🪟 聚焦至視窗：{win.title}")
    time.sleep(0.5)

WRONG_FORMAT_PIC = "wrong_format_2.png"
def save_pdf_gui(pdf_path):
    pyautogui.FAILSAFE = True
    time.sleep(2)
    try:
        wrong_format_pos = pyautogui.locateCenterOnScreen(WRONG_FORMAT_PIC, confidence=0.7)
    except Exception as e:
        logging.warning(f"⚠️ 搜尋 wrong_format.png 發生錯誤：{e}")
        wrong_format_pos = None   

    if wrong_format_pos is not None:
        wait_and_click(WRONG_FORMAT_PIC)
        time.sleep(1)
        wait_and_click("pdf_in_bar_options.png")
        time.sleep(1)
        wait_and_click("save_button.png", confidence=0.7)
        time.sleep(1)
    else:
        try:
            logging.info("🔄 嘗試點擊 save_button_with_frame.png")
            wait_and_click("save_button_with_frame.png", confidence=0.7, timeout=4)
            logging.info("🖱️ 成功點擊 save_button_with_frame.png")
        except Exception as e1:
            logging.warning("⚠️ 找不到 save_button_with_frame.png，嘗試改用 save_button.png")
            try:
                wait_and_click("save_button.png", confidence=0.7, timeout=4)
                logging.info("🖱️ 成功點擊 save_button.png")
            except Exception as e2:
                pyautogui.screenshot(f"save_button_failed_{int(time.time())}.png")
                raise Exception("❌ 兩種儲存按鈕都找不到，儲存失敗")

    time.sleep(2)


    # 等待 Windows 另存對話框出現
    for _ in range(10):
        if pyautogui.locateCenterOnScreen('win_save.png', confidence=0.7):
            break
        time.sleep(0.5)
    else:
        logging.warning("⚠️ 未偵測到另存對話框，仍繼續輸入路徑")

    # 貼上檔名路徑
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.3)
    pyperclip.copy(pdf_path)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyautogui.press('enter')
    logging.info("📂 已送出儲存")

    # 確認是否成功儲存
    for _ in range(20):
        if os.path.exists(pdf_path):
            logging.info("✅ PDF 儲存成功")
            return True
        time.sleep(0.5)
    
    return False



def trigger_print(driver):
    try:
        wait = WebDriverWait(driver, 10)
        # dropdown_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
            # 'body > div.wrapper.l-a4 > div.header.priority-nav.priority-nav-has-dropdown > span > button')))
        # driver.execute_script("arguments[0].click();", dropdown_btn)for attempt in range(3):  # 嘗試3次
        for attempt in range(3):
            try:
                dropdown_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                    'body > div.wrapper.l-a4 > div.header.priority-nav.priority-nav-has-dropdown > span > button')))
                dropdown_btn.click()
                break
            except Exception as e:
                logging.error(f"點擊失敗，錯誤：{e}")
                time.sleep(1)  # 等待再試
        time.sleep(1)
        # print_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[data-speed-action="print"]')))
        # driver.execute_script("arguments[0].click();", print_btn)
        # time.sleep(2)
        for attempt in range(3):
            try:
                print_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'li.sg-btn.sg-btn-default a[data-speed-action="print"]')))
                print_btn.click()
                break
            except Exception as e:
                logging.error(f"點擊失敗，錯誤：{e}")
                time.sleep(1)  # 等待再試
        time.sleep(2)

        # center_x = 1450
        # center_y = 885
 

        # screen_width, screen_height = pyautogui.size()

        # if 0 < center_x < screen_width and 0 < center_y < screen_height:
        #     time.sleep(1)
        #     pyautogui.moveTo(center_x, center_y, duration=1)
        #     pyautogui.click()
        #     logging.info("📤 已送出列印預覽要求（已用 pyautogui 點擊）")
        #     time.sleep(1)
        # else:
        #     raise ValueError(f"⚠️ Unsafe coordinates: ({center_x}, {center_y}) — aborting click.")
        wait_and_click("first_sure.png", confidence=0.9, timeout=10)
        logging.info("📤 已送出列印預覽要求")
        screen_width, screen_height = pyautogui.size()
        center_x = int(screen_width / 2)
        center_y = int(screen_height / 2)
        pyautogui.click(center_x, center_y)
        logging.info("🖱️ 點擊預覽區中央以聚焦")
        time.sleep(2)
        
    except Exception as e:
        logging.error(f"❌ trigger_print 發生例外：{e}")
        raise e


def retry_save_pdf_gui(driver, pdf_path, row_index, max_retries=1):
    for attempt in range(1, max_retries + 1):
        logging.info(f"🔁 嘗試第 {attempt} 次保存 PDF...")

        try:
            focus_latest_chrome_window()  # ⬅️ 強制聚焦
            trigger_print(driver)
            logging.info("✅ trigger_print 成功")

            if save_pdf_gui(pdf_path):
                logging.info("✅ PDF 儲存成功")
                time.sleep(1.5)
                return True
            else:
                logging.error("❌ PDF 儲存失敗（無例外但未成功）")

        except Exception as e:
            logging.warning(f"⚠️ 嘗試列印或儲存時發生例外：{e}")


        if attempt < max_retries:
            logging.info(f"🔄 關閉 PDF 頁籤並重新打開第 {row_index + 1} 筆")

            if len(driver.window_handles) > 1:
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                time.sleep(1)

            try:
                rows = driver.find_elements(By.CSS_SELECTOR, "tr.k-master-row")
                if row_index < len(rows):
                    old_tabs = driver.window_handles
                    rows[row_index].find_element(By.CSS_SELECTOR, "a[data-speed-id='docNumber']").click()
                    time.sleep(1)
                    new_tabs = driver.window_handles
                    new_tab = [t for t in new_tabs if t not in old_tabs]
                    if new_tab:
                        driver.switch_to.window(new_tab[0])
                    else:
                        driver.switch_to.window(driver.window_handles[-1])
                else:
                    logging.error(f"⚠️ row_index {row_index} 超出範圍")
                    return False
            except Exception as e:
                logging.error(f"❌ 重新開啟 PDF 頁面失敗：{e}")
                return False
        else:
            logging.info("🚫 已達最大重試次數")

    return False