# Robocorp Imports
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF

# Other
import time
from datetime import datetime

## Functions
# Step 1: Open the Intranet Website
def open_the_intranet_website():
    # Sets the browser configuration
    browser.configure(
        headless=False, # Headless toggle
        slowmo=1000, # Slow down the time in milliseconds that the robot waits for between 2 actions
    )
    # Navigates to URL
    browser.goto("https://robotsparebinindustries.com/")


# Step 2: Log in to the Intranet
def log_in():
    page = browser.page()
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")
    

# Step 3: Download Excel File
def download_excel_file():
    # Instantiates the HTTP class
    http = HTTP()
    # target_file is set to the resources folder
    # overwrite is set to True to overwrite the file if it already exists
    http.download("https://robotsparebinindustries.com/SalesData.xlsx", target_file="resources/excel", overwrite=True)


# Step 4: Read through excel file and insert data into form
def fill_form_with_excel_data():
    # Instantiates the Files class
    excel = Files()
    # Reads the excel file as a table
    excel.open_workbook("resources/excel/SalesData.xlsx")
    worksheet = excel.read_worksheet_as_table("data", header=True) # Name has to match worksheet name
    excel.close_workbook()
    for row in worksheet: 
        fill_and_submit_sales_form(row)
           
           
# Step 5: Fill out form and submit sales data
def fill_and_submit_sales_form(row):
    page = browser.page()
    page.fill("#firstname", row["First Name"])
    page.fill("#lastname", row["Last Name"])
    page.select_option("#salestarget", str(row["Sales Target"]))
    page.fill("#salesresult", str(row["Sales"]))
    page.click("text=Submit")


# Step 6: Take a screenshot of page
def collect_results():
    # Get the current date and time
    current_date_and_time = datetime.now()
    # Format the date and time
    timestamp = current_date_and_time.strftime("%Y-%m-%d_%H-%M-%S")
    # Take a screenshot of the page
    page = browser.page()
    page.screenshot(path=f"resources/screenshots/sales_summary_{timestamp}.png")
    return timestamp


# Step 7: Export the data to a pdf file
def export_as_pdf(timestamp):
    page = browser.page()
    sales_results_table = page.locator("#sales-results").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(sales_results_table, f"resources/pdf/sales_results_{timestamp}.pdf")


# Step 8: Log out
def log_out():
    page = browser.page()
    page.click("text=Log out")
