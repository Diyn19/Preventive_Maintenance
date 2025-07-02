from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import os
import time
import shutil

# 動態產生報表名稱
month_tag = datetime.today().strftime("%Y%m")
report_name = f"{month_tag}_HL_Maintain_Report.xlsx"

# 設定下載目錄
download_dir = os.path.abspath("downloads")  # 可自訂，這裡用同資料夾下的 downloads 資料夾
final_path = r"D:\SynologyDrive\flask\IM"

# 設定 Chrome 下載選項
options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

# 啟動瀏覽器
driver = webdriver.Chrome(options=options)
driver.get("http://192.168.1.252/service/")

# 登入
driver.find_element(By.NAME, "Account").send_keys("pos0800")
driver.find_element(By.NAME, "Password").send_keys("Pos0800")
driver.find_element(By.ID, "btn_login").click()

# 等待頁面跳轉並點選「POS服務工作統計表」
wait = WebDriverWait(driver, 10)
wait.until(EC.presence_of_element_located((By.LINK_TEXT, "服務資料查詢"))).click()
wait.until(EC.presence_of_element_located((By.LINK_TEXT, "POS服務工作統計表"))).click()

# 查詢條件設定
wait.until(EC.presence_of_element_located((By.NAME, "txtStartDate"))).clear()
wait.until(EC.presence_of_element_located((By.NAME, "txtEndDate"))).clear()
today = datetime.today().strftime("%Y/%m/%d")
driver.find_element(By.NAME, "txtStartDate").send_keys(today[:8] + "01")  # 當月1日
driver.find_element(By.NAME, "txtEndDate").send_keys(today)

driver.find_element(By.NAME, "selCustomer").send_keys("萊爾富")
driver.find_element(By.NAME, "selDepartment").send_keys("新北勤務一部")

# 點選查詢
driver.find_element(By.ID, "btn_query").click()

# 等表格出現後點擊匯出
wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "匯出成EXCEL"))).click()

# 等待檔案下載完成
print("等待檔案下載...")
downloaded = False
for i in range(30):  # 最多等 30 秒
    time.sleep(1)
    if os.path.exists(os.path.join(download_dir, report_name)):
        downloaded = True
        break

if downloaded:
    shutil.move(os.path.join(download_dir, report_name), os.path.join(final_path, report_name))
    print(f"✅ 成功下載並移動：{report_name}")
else:
    print("❌ 下載失敗或逾時")

driver.quit()
