from robocorp.tasks import task
from robocorp import browser

from RPA.FileSystem import FileSystem
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.Browser.Selenium import Selenium
from RPA.PDF import PDF
from RPA.Archive import Archive

import time

lib = Selenium()

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    open_browser()
    open_robot_order_website()
    # download_csv_file()
    orders = get_orders()
    for order in orders:
        close_annoying_modal()
        fill_the_form(order)
    archive_receipts()

def open_browser():
    """Open a browser"""
    lib.open_browser()

def open_robot_order_website():
    """Navigates to the given URL"""
    lib.go_to("https://robotsparebinindustries.com/#/robot-order")

def download_csv_file():
    """Download csv file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number","Head","Body","Legs","Address"]
    )
    return orders

def close_annoying_modal():
    """Close annoying modal/pop-up"""
    try:
        # Sesuaikan locator berdasarkan cara modal di-identifikasi pada halaman web Anda
        lib.click_button("xpath=//button[text()='OK']")
    except Exception as e:
        # Abaikan kesalahan jika modal tidak ditemukan
        pass

def fill_the_form(order):
    """"""
    lib.select_from_list_by_value("id=head", str(order["Head"]))
    body_option = str(order["Body"])
    if body_option == "1":
        lib.click_button("xpath=//label[contains(text(),'Roll-a-thor body')]/input")
    elif body_option == "2":
        lib.click_button("xpath=//label[contains(text(),'Peanut crusher body')]/input")
    elif body_option == "3":
        lib.click_button("xpath=//label[contains(text(),'D.A.V.E body')]/input")
    elif body_option == "4":
        lib.click_button("xpath=//label[contains(text(),'Andy Roid body')]/input")
    elif body_option == "5":
        lib.click_button("xpath=//label[contains(text(),'Spanner mate body')]/input")
    else:
        lib.click_button("xpath=//label[contains(text(),'Drillbit 2000 body')]/input")
    
    lib.find_element('class:form-control')
    lib.input_text('class:form-control', order["Legs"])
    lib.input_text("id=address", str(order["Address"]))
    lib.click_button("xpath=//button[text()='Preview']")
    lib.click_button("xpath=//button[text()='Order']")
    time.sleep(3)
    try:
        store_receipt_as_pdf(order)
        lib.click_button("xpath=//button[text()='Order another robot']")
    except Exception as e:
        print(f"Error: {str(e)}")
        print("Retrying...")
        # Jika tombol tidak ditemukan, panggil kembali fungsi fill_the_form
        fill_the_form(order)

def store_receipt_as_pdf(order):
    """Store the order receipt as a PDF file"""
    receipt =lib.get_element_attribute(locator="receipt",attribute="innerHTML")
    pdf = PDF()
    filesystem = FileSystem()
    filesystem.create_directory("output/receipts")
    path = "output/receipts/order_"+order['Order number']+".pdf"
    pdf.html_to_pdf(receipt,path)
    screenshot_robot(order)

def screenshot_robot(order):
    """"""
    # Capture a full-page screenshot
    lib.screenshot(locator="robot-preview-image",filename="output/receipts/"+order['Order number']+".png")

def archive_receipts():
    """"""
    archive = Archive()
    archive.archive_folder_with_zip("output/receipts","receipts.zip") 