import streamlit as st
import os
import shutil
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ===============================
# 🔧 CONFIGURATION
# ===============================
SHAREPOINT_FOLDER = r"C:\Users\skulkarni\OneDrive - mhp-group.com\TR-AD Governance & PMT Solutions - Test Projektleads"
DOWNLOAD_WAIT_TIME = 5  # seconds to wait for file to download fully

# ===============================
# 🧩 HELPER FUNCTIONS
# ===============================
def wait_for_download_and_rename(new_filename):
    st.info(f"⏳ Waiting for file to download to: {SHAREPOINT_FOLDER}")
    time.sleep(DOWNLOAD_WAIT_TIME)  # wait for download to finish

    files = [os.path.join(SHAREPOINT_FOLDER, f) for f in os.listdir(SHAREPOINT_FOLDER)]
    if not files:
        st.warning("⚠️ No files found in folder!")
        return

    latest_file = max(files, key=os.path.getctime)
    base, ext = os.path.splitext(latest_file)
    new_base = os.path.join(SHAREPOINT_FOLDER, new_filename)
    new_path = new_base + ext

    counter = 1
    while os.path.exists(new_path):
        new_path = f"{new_base}{counter}{ext}"
        counter += 1

    os.rename(latest_file, new_path)
    st.success(f"✅ File renamed to: {new_path}")

def click_ok_in_any_frame(driver):
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    for index, frame in enumerate(frames):
        try:
            driver.switch_to.frame(frame)
            ok_button = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.ID, "DLG_VARIABLE_dlgBase_BTNOK"))
            )
            ok_button.click()
            driver.switch_to.default_content()
            return True
        except Exception:
            driver.switch_to.default_content()
            continue
    try:
        ok_button = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.ID, "DLG_VARIABLE_dlgBase_BTNOK"))
        )
        ok_button.click()
        return True
    except Exception:
        return False

def click_excel_in_any_frame(driver):
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    for index, frame in enumerate(frames):
        try:
            driver.switch_to.frame(frame)
            excel_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.ID, "BUTTON_GROUP_ITEM_1_btn1_acButton"))
            )
            excel_button.click()
            driver.switch_to.default_content()
            return True
        except Exception:
            driver.switch_to.default_content()
            continue
    try:
        excel_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "BUTTON_GROUP_ITEM_1_btn1_acButton"))
        )
        excel_button.click()
        return True
    except Exception:
        return False

def run_sap_task(driver, url, file_name):
    st.info(f"🌐 Starting SAP automation for URL: {url}")
    driver.get(url)

    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, "iframe"))
        )
        clicked = click_ok_in_any_frame(driver)
    except TimeoutException:
        st.error("❌ Page or frames took too long to load.")
        clicked = False
    finally:
        driver.switch_to.default_content()

    if clicked:
        clicked_excel = click_excel_in_any_frame(driver)
        if clicked_excel:
            wait_for_download_and_rename(file_name)
        else:
            st.warning("⚠️ Could not click Excel button automatically.")

def clear_sharepoint_folder():
    if os.path.exists(SHAREPOINT_FOLDER):
        for filename in os.listdir(SHAREPOINT_FOLDER):
            file_path = os.path.join(SHAREPOINT_FOLDER, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.remove(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                st.warning(f"⚠️ Failed to delete {file_path}: {e}")

# ===============================
# 🚀 STREAMLIT INTERFACE
# ===============================
st.title("SAP PL Stunden Automation")

if st.button("PL Stunden"):
    st.info("Starting the automation...")

    clear_sharepoint_folder()

    # Edge setup
    options = webdriver.EdgeOptions()
    options.add_experimental_option("prefs", {
        "download.default_directory": SHAREPOINT_FOLDER,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })

    driver_path = os.path.join(os.path.dirname(__file__), "msedgedriver.exe")
    driver = webdriver.Edge(executable_path=driver_path, options=options)


    try:
        url1 = "https://bcsw-sap111.mymhp.net/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=00A9WIOTPIOBE954UD30UWXRD"
        url2 = "https://bcsw-sap111.mymhp.net/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=00A9WIOTPIOBE954UVIP3PLNT"

        run_sap_task(driver, url1, "M1_WEB_PL_PROJ_01A_all")

        driver.switch_to.new_window('tab')
        time.sleep(1)

        run_sap_task(driver, url2, "M1_WEB_PL_PROJ_01A_monat")

        st.success("✅ Both runs completed successfully.")

    finally:
        time.sleep(5)
        driver.quit()
        st.info("🚪 Browser closed.")
