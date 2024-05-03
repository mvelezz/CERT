from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive


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
        slowmo=1000,
    )
    open_robot_order_website()
    close_pop_up()
    fill_orders()
    collect_results()
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
        "orders.csv", columns=["Head", "Body", "Legs", "Address"]
    )

    return orders

def fill_the_form(robot):
    page = browser.page()

    page.select_option('#head', str(robot["Head"]))
    page.click('//*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div[{0}]/label'.format(robot["Body"]))
    page.fill("input[placeholder='Enter the part number for the legs']", robot["Legs"])
    page.fill("#address", robot["Address"])
    page.click("button:text('Preview')")

    while True:
        page.click("button:text('Order')")
        #handle_alert_popup()
        order_another_bot()
        close_pop_up()
    # path = store_reciept_PDF(int(robot["Order number"]))
    # screenshot = screenshot_robot(int(robot["Order number"]))
        #embed_screenshot_to_receipt(screenshot, path)


def fill_orders():
    orders = get_orders()

    for row in orders:
        fill_the_form(row)
    

def close_pop_up():
    page = browser.page()
    page.click("button:text('OK')")

def collect_results():
    """Take a screenshot of the page"""
    page = browser.page()
    page.screenshot(path="output/robot.png")

def store_reciept_PDF(order_num):
    page = browser.page()
    order_HTML = page.locator("receipt").inner_html()
    pdf = PDF()
    path = "output/receipt/{0}.pdf".format(order_num)
    pdf.html_to_pdf(order_HTML, path)
    return path

def screenshot_robot(order_number):
    """Takes screenshot of the ordered bot image"""
    page = browser.page()
    screenshot_path = "output/screenshots/{0}.png".format(order_number)
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_to_receipt(screenshot_path, pdf_path):
    """Embeds the screenshot to the bot receipt"""
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot_path, 
                                   source_path=pdf_path, 
                                   output_path=pdf_path)
    
def order_another_bot():
    """Clicks on order another button to order another bot"""
    page = browser.page()
    page.click("#order-another")


def handle_alert_popup():
    browser.handle_alert(action="ACCEPT")  # Accepts the alert
    # or
    browser.handle_alert(action="DISMISS")  # Dismisses the alert

def archive_receipts():
    """Archives all the receipt pdfs into a single zip archive"""
    lib = Archive()
    lib.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")