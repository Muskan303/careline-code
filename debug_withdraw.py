"""Debug: find withdraw on leave list page"""
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

opts = webdriver.ChromeOptions()
opts.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)

driver.get("http://gcplcareline.girnarsoft.com/admin/user/user/access?e=monika.bidawat@girnarsoft.com")
time.sleep(4)

# Try common leave-related URLs
urls = [
    "/employee/leave",
    "/employee/leave/list",
    "/employee/leave/history",
    "/employee/attendance/leave",
    "/employee/leave/withdraw",
]
for suffix in urls:
    url = f"http://gcplcareline.girnarsoft.com{suffix}"
    driver.get(url)
    time.sleep(2)
    bt = driver.find_element(By.TAG_NAME,"body").text[:200]
    print(f"URL: {url}")
    print(f"  Body: {bt[:100]}")
    print(f"  Current URL: {driver.current_url}")
    print()

# Check sidebar/nav for leave withdraw link
driver.get("http://gcplcareline.girnarsoft.com/employee/attendance")
time.sleep(3)
print("=== All nav links ===")
for a in driver.find_elements(By.TAG_NAME, "a"):
    try:
        href = a.get_attribute("href") or ""
        txt  = a.text.strip()
        if ("leave" in href.lower() or "withdraw" in href.lower() or
            "withdraw" in txt.lower()):
            print(f"  href='{href}' text='{txt}'")
    except Exception: pass

# Check left sidebar icons
print("\n=== Sidebar items ===")
for el in driver.find_elements(By.XPATH, "//nav//* | //aside//* | //*[contains(@class,'sidebar')]//*"):
    try:
        href = el.get_attribute("href") or ""
        txt  = el.text.strip()
        if href and "employee" in href and el.is_displayed():
            print(f"  tag={el.tag_name} href='{href}' text='{txt[:30]}'")
    except Exception: pass

driver.quit()
