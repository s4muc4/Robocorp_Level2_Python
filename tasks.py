import logging
from robocorp.tasks import task
from robocorp import browser
from RPA.Robocorp.Storage import Storage
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables # To work with CSV files
from RPA.PDF import PDF
import zipfile
import os

@task
def order_robots_from_RobotSpareBin():
    browser.configure(slowmo=100)

    open_robot_order_website()
    close_annoying_modal()
    table = get_orders()
    fill_the_form(table)
    archive_receipts()
    
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """

def open_robot_order_website():
    storage = Storage()
    

    url = storage.get_text_asset("RobotSpareBin_URL")
    browser.goto(url + "#/robot-order")
    

def close_annoying_modal():
    page = browser.page()
    pop_up_exists = page.is_visible("button:text('OK')", timeout=3000)
    if pop_up_exists:
        page.click("button:text('OK')")
    


def get_orders():
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite=True)

    csv = Tables()
    table = csv.read_table_from_csv("orders.csv", header=True)
    return table


def fill_the_form(table):
    
    page = browser.page()



    for order in table:
            
        logging.info(order["Order number"])
        close_annoying_modal()
        page.select_option("#head", order["Head"])
        page.check("#id-body-"+ order["Body"])
        page.fill("input[placeholder='Enter the part number for the legs']", order["Legs"])
        page.fill("#address", order["Address"])
        page.click("#order")

        if not page.is_visible("#order-another", timeout=1000):
            page.click("#order")
        if not page.is_visible("#order-another", timeout=1000):
            page.click("#order")
        if not page.is_visible("#order-another", timeout=1000):
            page.click("#order")
            

        order_number = page.text_content("//*[@id='receipt']/p[1]")
        path_pdf = store_receipt_as_pdf(order_number)
        path_screenshots = screenshot_robot(order_number)
        embed_screenshot_to_receipt(path_screenshots, path_pdf)

        page.click("#order-another")

        

def store_receipt_as_pdf(order_number):

    page = browser.page()

    receipt = page.locator("#receipt").inner_html()
    pdf = PDF()
    path = "output/receipts/"+order_number+".pdf"
    pdf.html_to_pdf(receipt, path)
    return path

def screenshot_robot(order_number):
    page = browser.page()
    path = "output/receipts/screenshots/"+order_number+".png"
    page.locator("#robot-preview-image").screenshot(path=path)
    return path

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()

    pdf.add_watermark_image_to_pdf(image_path=screenshot, source_path=pdf_file, output_path=pdf_file)

def archive_receipts():

    folder_path = "output/receipts"
    output_filename = "output/receipts.zip"

    with zipfile.ZipFile(output_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                zipf.write(file_path, os.path.relpath(file_path, os.path.dirname(folder_path)))