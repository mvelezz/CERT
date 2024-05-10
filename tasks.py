from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from robocorp.log import exception, info
from RPA.Browser.Selenium import Selenium
import time
from playwright.sync_api import expect


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=1,
    )
    open_robot_order_website()
    close_pop_up()
    fill_orders()
    archive_receipts()
    
    

    


def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"]
    )

    return orders

def fill_the_form(robot):
    """Fills the form with specified CSV form"""
    page = browser.page()

    body = robot["Body"]
    bodyTXT = f"input[value = '{body}']"

    page.select_option('#head', robot["Head"])
    page.locator(bodyTXT).click()
    page.fill("input[placeholder='Enter the part number for the legs']", robot["Legs"])
    page.fill("#address", robot["Address"])
    page.click("button:text('Preview')")
    #page.click("button:text('Order')")

    
    while True:
        try:
            page.click("button:text('Order')")
            #Using query selector because if no elements match the selector, the return value resolves to null
            order_another = page.query_selector("#order-another")
            if order_another:
                pdf_path=export_as_pdf(int(robot["Order number"]))
                image_path=collect_results(int(robot["Order number"]))
                embed_pdf(image_path, pdf_path)
        except:
            order_another_bot()
            close_pop_up()
            break
    

def fill_orders():
    """Loops through the rows of csv file"""
    orders = get_orders()

    for row in orders:
        fill_the_form(row)
    

def close_pop_up():
    """Closes initial pop up to waive rights"""
    page = browser.page()
    page.click("button:text('OK')")

def collect_results(order_number):
    """Take a screenshot of the page"""
    page = browser.page()
    image_path = f"output/screenshots/{order_number}.png"
    page.locator("#robot-preview-image").screenshot(path=image_path)
    return image_path

    
def order_another_bot():
    """Clicks on order another button to order another bot"""
    page = browser.page()
    page.click("#order-another")

def export_as_pdf(order_number):
    """Export the data to a pdf file"""
    page = browser.page()
    receipt_results_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    path = f"output/receipts/{order_number}.pdf"
    pdf.html_to_pdf(receipt_results_html, path)
    return path

def embed_pdf(image_path, pdf_path):
    pdf = PDF()

    pdf.add_watermark_image_to_pdf(
        image_path,
        pdf_path,
        output_path=pdf_path
    )

def archive_receipts():
    """Archives all the receipt pdfs into a single zip archive"""
    lib = Archive()
    lib.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")
