from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import shutil
import glob

# 設定 Edge 下載路徑與最終檔名
download_path = r"D:\SynologyDrive\flask\IM"
final_filename = "HL_Maintain_Report.xlsx"
final_filepath = os.path.join(download_path, final_filename)

# Edge 選項
options = webdriver.EdgeOptions()
options.use_chromium = True
options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

try:
    # 啟動 Edge driver
    driver = webdriver.Edge(options=options)
    wait = WebDriverWait(driver, 20)

    # 開啟網站並登入
    driver.get("http://192.168.1.252/service/")
    wait.until(EC.presence_of_element_located((By.ID, "login_id"))).send_keys("pos0800")
    driver.find_element(By.ID, "login_pwd").send_keys("Pos0800")
    driver.find_element(By.CSS_SELECTOR, 'input[type="submit"][value="登入"]').click()

    # 切換至 iframe
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "iframe")))

    # 回主框架選單操作
    driver.switch_to.default_content()
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "服務資料查詢"))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "POS服務工作統計表"))).click()

    # 回 iframe 選單查詢
    driver.switch_to.frame("iframe")
    time.sleep(2)

    wait.until(EC.element_to_be_clickable((By.NAME, "customer"))).click()
    time.sleep(0.5)
    wait.until(EC.element_to_be_clickable((By.XPATH, '//select[@name="customer"]/option[contains(text(),"萊爾富")]'))).click()

    wait.until(EC.element_to_be_clickable((By.NAME, "dept_id"))).click()
    time.sleep(0.5)
    wait.until(EC.element_to_be_clickable((By.XPATH, '//select[@name="dept_id"]/option[contains(text(),"新北勤務一部")]'))).click()

    driver.find_element(By.XPATH, '//input[@type="submit" and @value="查詢"]').click()

    # 等待「匯出成EXCEL」出現
    wait.until(EC.element_to_be_clickable((By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]')))

    # 刪除既有目標檔案避免 (1).xlsx
    if os.path.exists(final_filepath):
        os.remove(final_filepath)

    # 點擊「匯出成EXCEL」
    driver.find_element(By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]').click()

    # 等待下載完成
    time.sleep(10)

    # 找出最新 .xlsx 檔案
    xlsx_files = glob.glob(os.path.join(download_path, "*.xlsx"))
    if not xlsx_files:
        raise FileNotFoundError("❌ 找不到任何下載的 Excel 檔案。")

    latest_file = max(xlsx_files, key=os.path.getctime)

    # 重新命名
    if os.path.exists(final_filepath):
        os.remove(final_filepath)
    shutil.move(latest_file, final_filepath)

    print(f"✅ 下載成功，已儲存為：{final_filepath}")

except Exception as e:
    print("❌ 執行失敗：", str(e))

finally:
    try:
        driver.quit()
    except:
        pass
