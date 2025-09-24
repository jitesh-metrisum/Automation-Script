import os
from playwright.sync_api import Playwright, sync_playwright

# 📂 Define download directory
DOWNLOAD_DIR = r"C:\Users\maa00\OneDrive - Hem Corporation Pvt Ltd\Dashboard Sources\Focus Files\GT Customer Statements"

def run(playwright: Playwright) -> None:
    # Ensure the folder exists
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    # Open login page
    page.goto("http://13.235.198.88/focus8w/?AspxAutoDetectCookieSupport=1#")

    # 🔐 Login
    page.wait_for_selector("#txtUsername", timeout=10000)
    page.fill("#txtUsername", "MIS Local")
    page.wait_for_timeout(500)

    print("🔐 Filling password...")
    page.wait_for_selector("#txtPassword", timeout=10000)
    page.fill("#txtPassword", "Mis@123")  # ⚠️ fill real password
    
    page.locator("#ddlCompany").select_option("180")
    page.get_by_role("button", name="Sign In").click()

# GT Customer Statements

    # 📂 Navigation
    page.get_by_role("link", name=" Financials").click()
    page.get_by_role("link", name="Receivable and Payable").click()
    page.get_by_role("link", name="Customer Detail ").click()
    page.get_by_role("link", name="Customer Statements").click()

    # ⚙️ Select parameters
    page.get_by_text("Sundry Debtors", exact=True).first.click()
    page.get_by_role("cell", name="Sundry Debtors - Domestic", exact=True).dblclick()
    page.get_by_role("cell", name="Area - General Trade", exact=True).click()
    page.get_by_role("row", name="123 Area - General Trade 122122 Customer", exact=True).get_by_label("").check()
    page.get_by_role("checkbox", name="Adjustment as on Today").check()


    # Layout + Output
    page.locator("#RITLayout_").select_option("1622")
    page.locator("#RITOutput_").select_option("3")  # Excel export

    # 💾 Trigger download
    with page.expect_download(timeout=0) as download_info:
        page.locator("#reportViewControls span", has_text="Ok").click()

    download = download_info.value
    file_path = os.path.join(DOWNLOAD_DIR, download.suggested_filename)
    download.save_as(file_path)

    print(f"✅ Report downloaded: {file_path}")

    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

# ------------------------- SETUP ------------------------- #
# MT Customer Statements
import os
from playwright.sync_api import Playwright, sync_playwright

# 📂 Directory where reports will be saved
DOWNLOAD_DIR = r"C:\Users\maa00\OneDrive - Hem Corporation Pvt Ltd\Dashboard Sources\Focus Files\MT Customer Statements"

def run(playwright: Playwright) -> None:
    # Ensure the folder exists
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    # Open login page
    page.goto("http://13.235.198.88/focus8w/")

    # 🔐 Login
    page.wait_for_selector("#txtUsername", timeout=10000)
    page.fill("#txtUsername", "MIS Local")
    page.wait_for_timeout(500)

    print("🔐 Filling password...")
    page.wait_for_selector("#txtPassword", timeout=10000)
    page.fill("#txtPassword", "Mis@123")  # ⚠️ fill real password
    
    page.locator("#ddlCompany").select_option("180")
    page.get_by_role("button", name="Sign In").click()

    # MT Customer Statements


    # 📂 Navigation
    page.get_by_role("link", name=" Financials").click()
    page.get_by_role("link", name="Receivable and Payable").click()
    page.get_by_role("link", name="Customer Detail ").click()
    page.get_by_role("link", name="Customer Statements").click()

    # ⚙️ Select parameters
    page.get_by_text("Sundry Debtors", exact=True).first.click()
    page.get_by_role("cell", name="Sundry Debtors - Domestic", exact=True).dblclick()
    page.get_by_role("cell", name="Area - Modern Trade", exact=True).click()
    page.get_by_role("row", name="124 Area - Modern Trade 122121 Customer", exact=True).get_by_label("").check()
    page.get_by_role("checkbox", name="Adjustment as on Today").check()

    # Layout + Output
    page.locator("#RITLayout_").select_option("1623")
    page.locator("#RITOutput_").select_option("3")  # Excel export

    # 💾 Trigger download
    with page.expect_download(timeout=0) as download_info:
        page.locator("#reportViewControls span", has_text="Ok").click()

    download = download_info.value
    file_path = os.path.join(DOWNLOAD_DIR, download.suggested_filename)
    download.save_as(file_path)

    print(f"✅ Report downloaded to: {file_path}")

    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

    # ------------------------- SETUP ------------------------- #
# Customer Ageing Details - MT
import os
from playwright.sync_api import Playwright, sync_playwright


def run(playwright: Playwright) -> None:
    # ✅ Download folder
    download_dir = r"C:\Users\maa00\OneDrive - Hem Corporation Pvt Ltd\Dashboard Sources\Focus Files\MT Outstanding Report"
    os.makedirs(download_dir, exist_ok=True)

    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    # Go to login page
    page.goto("http://13.235.198.88/focus8w/#")

    # Login
    print("🔐 Filling username...")
    page.wait_for_selector("#txtUsername", timeout=10000)
    page.fill("#txtUsername", "MIS Local")

    print("🔐 Filling password...")
    page.wait_for_selector("#txtPassword", timeout=10000)
    page.fill("#txtPassword", "Mis@123")  # ⚠️ real password
    page.locator("#ddlCompany").select_option("180")
    page.get_by_role("button", name="Sign In").click()

    # Navigate to Customer Ageing Details
    page.get_by_role("link", name=" Financials").click()
    page.get_by_role("link", name="Receivable and Payable").click()
    page.get_by_role("link", name="Customer Detail ").click()
    page.get_by_role("link", name="Customer Ageing Details").click()

    # Apply filters
    page.get_by_text("Sundry Debtors", exact=True).first.click()
    page.get_by_role("cell", name="Sundry Debtors - Domestic", exact=True).dblclick()
    page.get_by_role("cell", name="Area - Modern Trade", exact=True).click()
    page.get_by_role("row", name="124 Area - Modern Trade 122121 Customer", exact=True).get_by_label("").check()
    page.get_by_role("checkbox", name="Adjustment as on Today").check()
    page.wait_for_timeout(1000)

    # ✅ Select Excel/CSV output
    page.locator("#RITOutput_").select_option("3")
    page.wait_for_timeout(1000)
    page.select_option("#RITLayout_", value="1581")

    # ✅ Ok button triggers the download
    with page.expect_download() as download_info:
        page.locator("#reportViewControls").get_by_text("Ok").click()
    download = download_info.value

    # ✅ Save downloaded file
    save_path = os.path.join(download_dir, download.suggested_filename)
    download.save_as(save_path)
    print(f"✅ Downloaded file saved to: {save_path}")

    # Cleanup
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)



# -------------------------------------------------------------------------------------------------------------
#SALES REGISTER AND SALES REGISTER LINE WISE

import os
from datetime import datetime
from playwright.sync_api import sync_playwright

DOWNLOAD_DIR = r"C:\Users\maa00\OneDrive - Hem Corporation Pvt Ltd\Dashboard Sources\Focus Files\Sales Registers 25-26"

def run(playwright):
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    print("🌐 Opening login page...")
    page.goto("http://13.235.198.88/focus8w/", timeout=0)

    print("🔐 Logging in...")
    page.wait_for_selector("#txtUsername", timeout=10000)
    page.fill("#txtUsername", "MIS Local")

        # Step 2: Fill password
    print("🔐 Filling password...")
    page.wait_for_selector("#txtPassword", timeout=10000)
    page.fill("#txtPassword","Mis@123" )
    page.select_option("#ddlCompany", value="180")
    page.click("#btnSignin")

    print("📂 Navigating to Sales Register...")
    page.get_by_role("link", name=" Financials").click()
    print("Navigated to Financials")
    page.get_by_role("link", name="Reports ").click()
    print("Navigated to Reports")
    page.get_by_role("link", name="Sales Report ").click()
    print("Navigated to Sales Report")
    page.get_by_role("link", name="Sales Register").click()
    

    print("⚙️ Selecting options...")
    page.get_by_role("checkbox", name="Include Sales Return voucher").check()
    page.select_option("#RITOutput_", value="3")  # Excel

    print("💾 Clicking OK and waiting for download...")
    with page.expect_download(timeout=30 * 60 * 1000) as download_info:
        # 👇 Targeting only the report's Ok button
        page.locator("#reportViewControls span:has-text('Ok')").click()

    download = download_info.value
    file_path = os.path.join(DOWNLOAD_DIR, download.suggested_filename)
    download.save_as(file_path)

    print(f"✅ File downloaded and saved to: {file_path}")

    browser.close()

if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)





# ------------------------------------------------------------------------------------------------------------
# Register-Sales Line Wise



import os
from datetime import datetime
from playwright.sync_api import sync_playwright

DOWNLOAD_DIR = r"C:\Users\maa00\OneDrive - Hem Corporation Pvt Ltd\Dashboard Sources\Focus Files\Register sales line wise 25-26"

def run(playwright):
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    print("🌐 Opening login page...")
    page.goto("http://13.235.198.88/focus8w/?AspxAutoDetectCookieSupport=1", timeout=0)

    print("🔐 Logging in...")
    page.wait_for_selector("#txtUsername", timeout=10000)
    page.fill("#txtUsername", "MIS Local")

        # Step 2: Fill password
    print("🔐 Filling password...")
    page.wait_for_selector("#txtPassword", timeout=10000)
    page.fill("#txtPassword", "Mis@123")
    page.select_option("#ddlCompany", value="180")
    page.click("#btnSignin")

    print("📂 Navigating to Sales Register...")
    page.get_by_role("link", name=" Financials").click()
    page.get_by_role("link", name="Reports ").click()
    page.get_by_role("link", name="Sales Report ").click()
    page.get_by_role("link", name="Register-Sales Line Wise").click()

    print("⚙️ Selecting options...")
    page.select_option("#DateOptions_", value="9")  # Date Range
    page.select_option("#RITOutput_", value="3")  # Excel

    print("💾 Clicking OK and waiting for download...")
    with page.expect_download(timeout=30 * 60 * 1000) as download_info:
        # 👇 Targeting only the report's Ok button
        page.locator("#reportViewControls span:has-text('Ok')").click()

    download = download_info.value
    file_path = os.path.join(DOWNLOAD_DIR, download.suggested_filename)
    download.save_as(file_path)

    print(f"✅ File downloaded and saved to: {file_path}")

    browser.close()

if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)






# -------------------------------------------------------------------------------------------------------------------

