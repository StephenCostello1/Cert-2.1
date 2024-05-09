from robocorp.tasks import task
from robocorp import browser
import pandas as pd
import requests
from io import StringIO
import time
from RPA.PDF import PDF  # Import the PDF library from RPA Framework
from fpdf import FPDF  # Ensure FPDF is imported
from pathlib import Path  # Import Path from pathlib
import shutil

@task
def order_robots_from_RobotSpareBin():
    browser.configure(
        slowmo=100,
    )
    open_the_intranet_website()
    close_annoying_modal()
    fill_forms_with_order_data()
    archive_order_receipts()

def open_the_intranet_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def close_annoying_modal():
    page = browser.page()
    page.click("//button[normalize-space()='OK']")

def fill_forms_with_order_data():
    url = "https://robotsparebinindustries.com/orders.csv"
    headers = {'User-Agent': 'Mozilla/5.0'}
    
    response = requests.get(url, headers=headers)
    response.raise_for_status()  # This will raise an error if the request failed
    csv_data = StringIO(response.content.decode('utf-8'))
    orders = pd.read_csv(csv_data)
   
    for index, order_row in orders.iterrows():
        fill_and_submit_order_form(order_row)

def fill_and_submit_order_form(order_row):
    page = browser.page() 
    head_index = int(order_row['Head']) 
    head_value = str(head_index)  
    page.select_option("#head", value=head_value)
    body_index = int(order_row['Body'])
    body_selector = f"//input[@id='id-body-{body_index}']"
    page.click(body_selector)
    
    legs_value = order_row['Legs'] 
    placeholder_selector = "input[placeholder='Enter the part number for the legs']"
    page.fill(placeholder_selector, str(legs_value))

    address_value = order_row['Address'] 
    page.fill("#address", str(address_value))

    time.sleep(1)
    page.click("id=order")
    time.sleep(1)
    if page.is_visible("#order", timeout=1000):
        page.click("#order")
    if page.is_visible("#order", timeout=1000):
        page.click("#order")
    if page.is_visible("#order", timeout=1000):
        page.click("#order")
    if page.is_visible("#order", timeout=1000):
        page.click("#order")
    if page.is_visible("#order", timeout=1000):
        page.click("#order")
    time.sleep(1)
    order_number = read_order_number()
    take_screenshot_and_save_as_pdf(order_number)  # Call the new function here
    time.sleep(1)
    page.click("id=order-another")
    time.sleep(1)
    page.click("//button[normalize-space()='OK']")

def read_order_number():
    page = browser.page()
    time.sleep(1)  # Allow page to load
    if page.is_visible(".badge.badge-success"):
        order_number = page.text_content(".badge.badge-success")
        if order_number:
            trimmed_order_number = order_number[15:].strip()
            print(trimmed_order_number)
            return trimmed_order_number  # Return the order number for further use
    return None

def take_screenshot_and_save_as_pdf(order_number):
    page = browser.page()
    output_dir = Path("output/receipts")
    output_dir.mkdir(parents=True, exist_ok=True)  # Ensure the directory exists

    # Define paths for the screenshot and the resulting PDF
    screenshot_path = output_dir / f"{order_number}_screenshot.png"
    pdf_path = output_dir / f"{order_number}.pdf"

    # Take the screenshot
    page.screenshot(path=str(screenshot_path))

    # Create PDF and add the image
    pdf = FPDF()
    pdf.add_page()
    pdf.image(str(screenshot_path), x=10, y=8, w=180)  # Adjust dimensions as needed
    pdf.output(str(pdf_path))
    print(f"PDF saved at: {pdf_path}")
    screenshot_path.unlink()  # Remove the screenshot after embedding it into the PDF

def archive_order_receipts():
    # Define the directory containing the PDF receipts
    receipts_dir = Path("output/receipts")
    
    # Define the path for the output ZIP file (without the .zip extension)
    archive_path = Path("output/order_receipts_archive")
    
    # Create a ZIP archive containing all PDF receipts
    shutil.make_archive(archive_path, 'zip', receipts_dir)
    
    print(f"Created ZIP archive at: {archive_path}.zip")