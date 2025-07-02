from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import shutil
import glob
from datetime import datetime

# 當前年月
yyyymm = datetime.now().strftime("%Y%m")
download_path = r"D:\flask\IM"
base_filename = f"{yyyymm}_HL_Maintain_Report.xlsx"
pattern = os.path.join(download_path, f"{yyyymm}_HL_Maintain_Report*.xlsx")

# Edge 選項
options = webdriver.EdgeOptions()
options.use_chromium = True
options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

def safe_wait_and_find(driver, by, value, wait_time=10, retries=1, refresh_if_fail=False):
    """
    嘗試等待特定元素出現，若失敗可重新整理並再試
    """
    for attempt in range(retries + 1):
        try:
            wait = WebDriverWait(driver, wait_time)
            return wait.until(EC.presence_of_element_located((by, value)))
        except:
            if refresh_if_fail and attempt < retries:
                print(f"⚠️ 找不到元素 {value}，重新整理頁面 (第 {attempt+1} 次)...")
                driver.refresh()
                time.sleep(3)
            else:
                raise

try:
    driver = webdriver.Edge(options=options)
    wait = WebDriverWait(driver, 20)

    # 開啟網站並登入
    driver.get("http://192.168.1.252/service/")
    safe_wait_and_find(driver, By.ID, "login_id").send_keys("pos0800")
    driver.find_element(By.ID, "login_pwd").send_keys("Pos0800")
    driver.find_element(By.CSS_SELECTOR, 'input[type="submit"][value="登入"]').click()

    # 切換至 iframe
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "iframe")))

    # 回主框架點擊選單（支援自動重試）
    driver.switch_to.default_content()
    safe_wait_and_find(driver, By.LINK_TEXT, "服務資料查詢", retries=1, refresh_if_fail=True).click()
    time.sleep(1)
    safe_wait_and_find(driver, By.LINK_TEXT, "POS服務工作統計表", retries=1, refresh_if_fail=True).click()

    # 切回 iframe
    driver.switch_to.frame("iframe")
    time.sleep(1)

    # 選擇「萊爾富」
    customer_field = safe_wait_and_find(driver, By.NAME, "customer", retries=1, refresh_if_fail=True)
    customer_field.click()
    time.sleep(0.5)
    safe_wait_and_find(driver, By.XPATH, '//select[@name="customer"]/option[contains(text(),"萊爾富")]').click()

    # 選擇「新北勤務一部」
    dept_field = safe_wait_and_find(driver, By.NAME, "dept_id")
    dept_field.click()
    time.sleep(0.5)
    safe_wait_and_find(driver, By.XPATH, '//select[@name="dept_id"]/option[contains(text(),"新北勤務一部")]').click()

    # 按下查詢
    driver.find_element(By.XPATH, '//input[@type="submit" and @value="查詢"]').click()

    # 等待匯出按鈕（支援重試）
    safe_wait_and_find(driver, By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]', retries=1, refresh_if_fail=True)

    # 刪除舊檔案
    for f in glob.glob(pattern):
        os.remove(f)

    # 匯出
    driver.find_element(By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]').click()

    # 等待下載完成
    max_wait = 20
    downloaded_file = None
    for _ in range(max_wait):
        files = glob.glob(pattern)
        if files:
            downloaded_file = max(files, key=os.path.getctime)
            break
        time.sleep(1)

    driver.quit()

    if downloaded_file and os.path.exists(downloaded_file):
        print(f"✅ 下載完成：{os.path.basename(downloaded_file)}")
    else:
        print("❌ 錯誤：未找到下載的檔案。")

except Exception as e:
    print(f"❌ 發生錯誤：{e}")
    if 'driver' in locals():
        driver.quit()
