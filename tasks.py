import os
import zipfile
from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from PIL import Image
from RPA.FileSystem import FileSystem
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
        slowmo=50,
    )
    open_robot_order_website()
    close_annoying_modal()
    orders = get_orders()
    for order in orders:
        screenshot_filename = fill_the_form(order)
        pdf_receipt = store_receipt_as_pdf(screenshot_filename)
        embed_screenshot_to_receipt(screenshot_filename, pdf_receipt, order["Order number"])
    archive_receipts()


def open_robot_order_website():
    """Open the robot order website"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def close_annoying_modal():
    """Close pop-up"""
    page = browser.page()
    page.click("button:text('OK')")


def get_orders():
    """
    Download orders file
    Read it as a table
    Fill the form
    """
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite=True)
    tables = Tables()
    orders_data = tables.read_table_from_csv(path="orders.csv", header=True)
    return orders_data


def fill_the_form(order):
    """Fills the page order"""
    page = browser.page()
    #head
    page.click("#head")
    page.select_option("#head", value=(order["Head"]))
    #body
    page.click(f"[name='body'][value='{order['Body']}']")
    #legs
    page.fill("[placeholder='Enter the part number for the legs']", value=str(order["Legs"]))
    #address
    page.fill("#address", str(order["Address"]))
    #preview and screenshot
    page.click("#preview")
    robot_preview = page.locator("#robot-preview")
    order_number = str(order["Order number"])

    # Assicurati che la cartella esista prima di salvarvi l'immagine
    screenshot_dir = "output/receipt-order"
    if not os.path.exists(screenshot_dir):
        os.makedirs(screenshot_dir)
    # Screenshot della preview del robot
    screenshot_filename = f"{screenshot_dir}/robot-preview-image-order-{order_number}.png"
    print(f"Saving ROBOT-PREVIEW screenshot at {screenshot_filename}")
    robot_preview.screenshot(path=screenshot_filename)
    #confirm order
    page.click("#order")
    #check if there is an error and click again if persist
    message_error = page.locator(".alert.alert-danger")
    while message_error.is_visible():
        page.click("#order")
        message_error = page.locator(".alert.alert-danger")
    
    return screenshot_filename


def store_receipt_as_pdf(order_number):
    """Store receipt as PDF"""
    try:
        # Screenshot della ricevuta
        page = browser.page()
        screenshot_receipt_dir = "output/receipt-order"
        receipt_file_name = f"{screenshot_receipt_dir}/receipt-order-{order_number}.png"
        receipt = page.locator("#receipt")    
        screenshot_receipt = receipt.screenshot(path=receipt_file_name)
        print(f"Saving RECEIPT screenshot at {receipt_file_name}")
        
        # Usa Pillow per aprire il file PNG
        img = Image.open(receipt_file_name)
        
        # Converti come PDF
        pdf_receipt = receipt_file_name.replace('.png', '.pdf')
        img.convert("RGB").save(pdf_receipt)
                
        # Aggiungi il file immagine al PDF
        pdf = PDF()
        pdf.add_files_to_pdf(files={pdf_receipt}, target_document=pdf_receipt)
        print(f"RECEIPT converted to PDF!")
        
        # Rimuovi il file PNG
        #fs = FileSystem()
        #fs.remove_file(path=screenshot_receipt)
        
        return pdf_receipt

    except Exception as e:
        print(f"Errore durante la creazione del PDF: {e}")
        raise


def embed_screenshot_to_receipt(screenshot_filename, receipt_pdf, order_number):
    """Embed the screenshot to the receipt PDF"""
    try:
        # Verifica se il file immagine esiste
        if not os.path.exists(screenshot_filename):
            raise FileNotFoundError(f"Immagine non trovata: {screenshot_filename}")

        # Converti PNG in PDF
        img = Image.open(screenshot_filename)
        pdf_robot_from_file_img = screenshot_filename.replace('.png', '.pdf')
        img.convert("RGB").save(pdf_robot_from_file_img)        # Screenshot del robot trasformato in PDF
        
        # Unisci il PDF della ricevuta con il PDF dell'immagine del robot
        pdf = PDF()
        files_to_merge = [receipt_pdf, pdf_robot_from_file_img]
        final_pdf_name = f"output/receipt-order/receipt-order-{order_number}-complete.pdf"
        pdf.add_files_to_pdf(files=files_to_merge, target_document=final_pdf_name, append=True)

        # Rimuovi i file temporanei
        fs = FileSystem()
        fs.remove_file(path=screenshot_filename)
        fs.remove_file(path=pdf_robot_from_file_img)
        
        # Clicca su "order another"
        page = browser.page()
        page.click("#order-another")
        close_annoying_modal()

    except Exception as e:
        print(f"Errore durante l'integrazione del PDF: {e}")
        raise


def archive_receipts():
    """ Crea archivio .ZIP contenente tutti gli ordini processati """
    receipts_dir = "output/receipt-order"
    zip_dir = "output"

    # Verifica che la directory output esista
    if not os.path.exists(zip_dir):
        os.makedirs(zip_dir)
    
    # Ottieni tutti i file .pdf dalla directory receipts_dir
    receipts_files = [file for file in os.listdir(receipts_dir) if file.endswith('.pdf')]
    
    # Verifica che ci siano file .pdf
    if not receipts_files:
        print(f"Nessun file PDF trovato nella directory {receipts_dir}.")
        return

    # Creazione dell'archivio ZIP
    archive_name = os.path.join(zip_dir, "receipts.zip")  # Specifica dove vuoi salvare l'archivio ZIP
    print(f"Creando l'archivio ZIP: {archive_name}")

    try:
        # Usa il contesto di gestione con zipfile per evitare problemi di apertura del file
        with zipfile.ZipFile(archive_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file in receipts_files:
                file_path = os.path.join(receipts_dir, file)
                print(f"Aggiungendo {file} all'archivio {archive_name}")
                zipf.write(file_path, os.path.basename(file))  # Aggiungi il file al ZIP con il nome base
                
        print(f"Archivio ZIP creato con successo: {archive_name}")
        return archive_name

    except Exception as e:
        print(f"Errore durante la creazione dell'archivio ZIP: {e}")
        return None


