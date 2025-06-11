import os
import csv
import time
import zipfile
from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
from PIL import Image, ImageDraw, ImageFont
import logging
from datetime import datetime

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
fs = FileSystem()

@task
def create_robot():
    """Creates a robot. *TASK*"""
    try:
        set_browser_time_delay(100)
        open_robot_website()
        download_csv_file()
        with open("input/orders.csv") as file:
            for row in csv.DictReader(file):
                order_number = row["Order number"]
                fill_form(row)
                create_output_directory()
                screenshot_robot_preview(order_number)
                submit_form(order_number)
                create_pdf(order_number)
        create_zip_file()
    finally:
        time.sleep(10)
        delete_robot_parts_folder()

def set_browser_time_delay(delay):
    """Sets the browser time delay."""
    browser.configure(
        slowmo=delay,
    )

def open_robot_website():
    """Opens the robot website and clicks the OK button."""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    browser.page().click("button:text('OK')")

def download_csv_file():
    """Downloads the excel file."""
    HTTP().download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True, target_file="input/orders.csv")

def screenshot_robot_preview(order_number):
    """Downloads and combines the robot part images into one image."""
    page = browser.page()
    page.locator("#robot-preview-image").screenshot(path=fs.join_path("output/robot_parts", f"robot_{order_number}.png"))

def fill_form(row):
    """Fills and submits the form."""
    page = browser.page()
    page.select_option("#head", row["Head"])
    page.click(f"#id-body-{row['Body']}")
    page.fill("input[type='number'][placeholder='Enter the part number for the legs']", row["Legs"])
    page.fill("#address", row["Address"])
    page.click("#preview")    

def submit_form(order_number):
    page = browser.page()
    max_attempts = 10
    attempt = 0
    
    while attempt < max_attempts:
        page.click("#order")
        if page.locator(".alert-danger").count() == 0:
            break
        attempt += 1
        page.wait_for_timeout(1000)
    
    if attempt == max_attempts:
        logger.error(f"Failed to submit the form after {max_attempts} attempts for order {order_number}")
        return

def create_combined_image(receipt_text, robot_path, order_id):
    """Creates a combined image with receipt and robot."""
    # Create receipt image
    receipt_img = Image.new('RGB', (400, 150), 'white')
    draw = ImageDraw.Draw(receipt_img)
    font = ImageFont.truetype(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ARIAL.TTF'), size=16)
    draw.text((30, 30), receipt_text, fill='black', font=font)
    receipt_path = fs.join_path("output/robot_parts", f"receipt_{order_id}.png")
    receipt_img.save(receipt_path, quality=100)
    
    # Resize robot image
    robot_img = Image.open(robot_path)
    # Make the image larger - 3/4 of original size
    new_width = int(robot_img.width * 0.6)
    new_height = int(robot_img.height * 0.6)
    robot_img = robot_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
    robot_img.save(robot_path, quality=100, optimize=False)
    
    # Create combined image
    combined_img = Image.new('RGB', (500, 600), 'white')
    combined_img.paste(receipt_img, (50, 0))  # Center receipt (500-400)/2 = 50
    robot_x = (500 - new_width) // 2 # Center Robot
    combined_img.paste(robot_img, (robot_x, 200))
    combined_path = fs.join_path("output/robot_parts", f"combined_{order_id}.png")
    combined_img.save(combined_path, quality=100)
    
    return combined_path

def get_receipt_text(page, order_id):
    """Creates the receipt text from the page content."""
    receipt_text = f"Order ID: {order_id}\n"
    receipt_text += f"Date: {page.locator('#receipt div').first.text_content()}\n"
    receipt_text += f"Address: {page.locator('#receipt p').first.text_content()}\n"
    receipt_text += f"Head: {page.locator('#parts div').nth(0).text_content().replace('Head: ', '')}\n"
    receipt_text += f"Body: {page.locator('#parts div').nth(1).text_content().replace('Body: ', ''),}\n"
    receipt_text += f"Legs: {page.locator('#parts div').nth(2).text_content().replace('Legs: ', '')}\n"
    return receipt_text

def create_pdf(order_number):
    """Creates a PDF with the receipt text and robot image."""
    pdf = PDF()
    page = browser.page()
    order_id = page.locator('#receipt .badge').text_content()
    pdf_path = fs.join_path("output/robot_parts", f"order_{order_id}.pdf")
    
    receipt_text = get_receipt_text(page, order_id)
    robot_path = fs.join_path("output/robot_parts", f"robot_{order_number}.png")
    combined_path = create_combined_image(receipt_text, robot_path, order_id)

    pdf.add_files_to_pdf(
        files=[combined_path],
        target_document=pdf_path,
        append=False
    )
    page.click("#order-another")
    page.click("button:text('OK')")

def create_zip_file():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_path = fs.join_path("output", f"robot_orders_{timestamp}.zip")
    
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file in fs.list_files_in_directory("output/robot_parts"):
            if file.name.endswith('.pdf'):
                file_path = fs.join_path("output/robot_parts", file.name)
                zipf.write(file_path, file.name)
    
    logger.info(f"Created zip file with all orders: {zip_path}")

def delete_robot_parts_folder():
    folder_path = "output/robot_parts"
    
    # Delete all files in the folder
    for file in fs.list_files_in_directory(folder_path):
        try:
            fs.remove_file(fs.join_path(folder_path, file.name))
        except Exception as e:
            logger.error(f"Error deleting file {file.name}: {e}")
    
    # Remove the empty directory
    try:
        fs.remove_directory(folder_path, recursive=True)
    except Exception as e:
        logger.error(f"Error removing directory: {e}")

def create_output_directory():
    """Creates the output directory."""
    fs.create_directory("output/robot_parts", exist_ok=True)