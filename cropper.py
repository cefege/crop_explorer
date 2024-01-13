import pandas as pd
import os
from playwright.sync_api import sync_playwright
from fpdf import FPDF
import time
from PIL import Image
from datetime import datetime
from PyPDF2 import PdfMerger
from tqdm import tqdm
import re


# Get the current date in USA format (mm-dd-yyyy) and save it as a global variable
date = datetime.now().strftime("%m-%d-%Y")


####EMAIL SETUP SECTION####
smtp_server = "smtp.example.com"  # Replace with your SMTP server
port = 465  # Replace with your SMTP server's SSL port number
sender_email = "your-email@example.com"  # Replace with your email address
sender_password = "your-password"  # Replace with your email password
recipient_emails = [
    "recipient1@example.com",
    "recipient2@example.com",
]  # Replace with recipient email addresses
subject = "Crop Report"
body = "Please find attached the latest crop report."
attachment_path = os.path.join("output/", f"crop_report_{date}.pdf")
####EMAIL SETUP ####


# Define the URL of the Google Sheet in CSV export format
sheet_url = "https://docs.google.com/spreadsheets/d/1LOdsJnAssCGakm-lOd0zoOnuWClR0WKKuvwi0xD6KIQ/export?format=csv"

# Global variable to define wait time between rows
wait_time_between_rows = 0


# Update images folder to include the date
images_base_dir = f"images/{date}/"


def read_csv_from_url(url):
    # Read the CSV data into a DataFrame
    df = pd.read_csv(url)

    # Replace the Unicode character '\u2019' with an ASCII apostrophe "'" in all string columns
    df = df.applymap(lambda x: x.replace("\u2019", "'") if isinstance(x, str) else x)

    return df


# Get the DataFrame from the Google Sheets CSV export URL
df = read_csv_from_url(sheet_url)


def get_image_size(image_path):
    with Image.open(image_path) as img:
        return img.size


def fit_image_size(image_size, max_width, max_height):
    width, height = image_size
    aspect_ratio = width / height
    if width > max_width:
        width = max_width
        height = width / aspect_ratio
    if height > max_height:
        height = max_height
        width = height * aspect_ratio
    return width, height


def create_pdf(images_dir, image_files, commodity, location, small_images=True):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Set title format
    title = f"Commodity: {commodity}, Location: {location}"
    pdf.set_font("Arial", size=12)
    pdf.cell(0, 5, title, ln=True, align="C")

    # Define maximum dimensions for images to fit on a page
    max_width = (pdf.w - 30) / (
        2 if small_images else 1
    )  # Two images per row if small_images is True
    max_height = pdf.h - 20  # Adjusted for the title height

    # Initialize variables to track image placement
    x_offset = 10
    y_offset = 15  # Adjusted for the title height
    row_height = 0

    for image_file in image_files:
        image_path = os.path.join(images_dir, image_file)
        img_width, img_height = get_image_size(image_path)

        # Scale image if it's larger than half the page width
        if img_width > max_width or img_height > max_height:
            img_width, img_height = fit_image_size(
                (img_width, img_height), max_width, max_height
            )

        # Check if the image fits in the current row, otherwise start a new row
        if small_images and x_offset + img_width > pdf.w - 10:
            x_offset = 10
            y_offset += row_height + 10
            row_height = 0

        # Add a new page if the image does not fit on the current page
        if y_offset + img_height > pdf.h - 20:
            pdf.add_page()
            # Add title to the new page
            pdf.cell(0, 10, title, ln=True, align="C")
            x_offset = 10
            y_offset = 30  # Reset offset for the title
            row_height = 0

        # Place the image on the page
        pdf.image(image_path, x=x_offset, y=y_offset, w=img_width, h=img_height)

        # Update the offset for the next image
        if small_images:
            x_offset += img_width + 10
        else:
            x_offset = 10
            y_offset += img_height + 10

        row_height = max(row_height, img_height)

    # Save the PDF to a file
    pdf.output(os.path.join(images_dir, f"{os.path.basename(images_dir)}.pdf"))
    # print(f"PDF created for {os.path.basename(images_dir)}")

    # Delete all ".png" images in the folder
    for image_file in image_files:
        if image_file.lower().endswith(".png"):
            os.remove(os.path.join(images_dir, image_file))


def get_and_save_images(url, row_number, commodity, location):
    images_dir = os.path.join(images_base_dir, f"row_{row_number}")
    pdf_files = (
        [f for f in os.listdir(images_dir) if f.endswith(".pdf")]
        if os.path.exists(images_dir)
        else []
    )

    # Check if PDF already exists in the row folder
    if pdf_files:
        # print(f"PDF already exists for row {row_number}, skipping...")
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        images_dir = os.path.join(images_base_dir, f"row_{row_number}")
        image_files = []

        try:
            # print(f"Navigating to the URL for row {row_number}...")
            page.goto(url)

            # Scroll to the bottom of the page
            # print("Scrolling to the bottom of the page...")
            page.evaluate("window.scrollTo(0, document.body.scrollHeight);")

            # Wait for the load state to be 'networkidle'
            # print("Waiting for the page to load content...")
            page.wait_for_load_state("networkidle")

            # Extract image elements
            # print("Extracting image elements...")
            image_elements = page.query_selector_all("img")

            # Create images directory if it doesn't exist
            if not os.path.exists(images_dir):
                os.makedirs(images_dir)

            # Save each image
            for i, image_element in enumerate(image_elements):
                # Check if the image URL matches the criteria
                src = image_element.get_attribute("src")
                if (
                    "/rssiws/images/" in src
                    and src.endswith(".gif")
                    or "ChartImg.ax" in src
                ):
                    # Take the buffer of the image
                    buffer = image_element.screenshot()

                    # Parse the name of the file from the URL
                    filename = f"image_{i}.png"
                    image_files.append(filename)

                    # Save the image to the images directory
                    file_path = os.path.join(images_dir, filename)
                    with open(file_path, "wb") as file:
                        file.write(buffer)
                        # print(f"Saved {filename}")

        except Exception as e:
            print(f"An error occurred: {e}")
        finally:
            browser.close()

        # After saving images, create a PDF with all images
        create_pdf(images_dir, image_files, commodity, location, small_images=True)


def get_and_save_special_screenshots(url, row_number, commodity, location):
    images_dir = os.path.join(images_base_dir, f"row_{row_number}")
    pdf_files = (
        [f for f in os.listdir(images_dir) if f.endswith(".pdf")]
        if os.path.exists(images_dir)
        else []
    )

    # Check if PDF already exists in the row folder
    if pdf_files:
        print(f"PDF already exists for row {row_number}, skipping...")
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        images_dir = os.path.join(images_base_dir, f"row_{row_number}")
        image_files = []

        try:
            # print(f"Navigating to the URL for row {row_number}...")
            page.goto(url)

            # Wait for the load state to be 'networkidle'
            print("Waiting for the page to load content...")
            page.wait_for_load_state("networkidle")

            # Create images directory if it doesn't exist
            if not os.path.exists(images_dir):
                os.makedirs(images_dir)

            # Screenshot of each "center" element
            # print("Taking screenshots of 'center' elements...")
            center_elements = page.query_selector_all("center")
            for i, center_element in enumerate(center_elements):
                center_screenshot_buffer = center_element.screenshot()
                center_filename = f"center_element_{i}.png"
                image_files.append(center_filename)
                center_file_path = os.path.join(images_dir, center_filename)
                with open(center_file_path, "wb") as file:
                    file.write(center_screenshot_buffer)
                    print(f"Saved {center_filename}")

        except Exception as e:
            print(f"An error occurred: {e}")
        finally:
            browser.close()

        # After saving screenshots, create a PDF with all images
        create_pdf(images_dir, image_files, commodity, location, small_images=False)


# Function to merge PDFs from row folders in alphabetical order
def merge_pdfs_in_row_folders(base_dir):
    # Initialize a PdfMerger object
    pdf_merger = PdfMerger()

    # Define a lambda function to extract the numerical part of the directory name and convert it to an integer
    natural_key = lambda s: int(re.search(r"row_(\d+)", s).group(1))

    # Get the list of row directories in the base directory and sort them using the natural_key function
    row_dirs = sorted(
        [d for d in os.listdir(base_dir) if d.startswith("row_")], key=natural_key
    )

    # Iterate over each row directory
    for row_dir in row_dirs:
        row_path = os.path.join(base_dir, row_dir)

        # Find the first PDF file in the current row directory
        pdf_files = sorted([f for f in os.listdir(row_path) if f.endswith(".pdf")])
        if pdf_files:
            first_pdf_file = pdf_files[0]
            pdf_file_path = os.path.join(row_path, first_pdf_file)

            # Append the PDF to the merger
            pdf_merger.append(pdf_file_path)
            print(f"Added {first_pdf_file} to the merged PDF.")

    # Define the output directory
    output_dir = "output"

    # Ensure the output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Output the merged PDF
    merged_pdf_path = os.path.join("output/", f"crop_report_{date}.pdf")
    pdf_merger.write(merged_pdf_path)
    pdf_merger.close()
    # print(f"Merged PDF created at {merged_pdf_path}")


def with_retry(func, *args, max_attempts=100, **kwargs):
    for attempt in range(max_attempts):
        try:
            # Call the function with the provided arguments
            func(*args, **kwargs)
            return  # If the function succeeds, exit the retry loop
        except Exception as e:
            print(f"Attempt {attempt + 1}: An error occurred - {e}")
            time.sleep(1)  # Wait for 1 second before the next attempt
    raise Exception(f"Function {func.__name__} failed after {max_attempts} attempts")


def send_email_with_attachment(
    smtp_server,
    port,
    sender_email,
    sender_password,
    recipient_emails,
    subject,
    body,
    attachment_path,
):
    # Create the email message
    msg = EmailMessage()
    msg["From"] = sender_email
    msg["To"] = ", ".join(recipient_emails)
    msg["Subject"] = subject
    msg.set_content(body)

    # Read the attachment and set the correct MIME type
    with open(attachment_path, "rb") as file:
        file_data = file.read()
        file_name = os.path.basename(attachment_path)

    msg.add_attachment(
        file_data, maintype="application", subtype="octet-stream", filename=file_name
    )

    # Connect to the SMTP server and send the email
    with smtplib.SMTP_SSL(smtp_server, port) as server:
        server.login(sender_email, sender_password)
        server.send_message(msg)

    print(f"Email sent to {', '.join(recipient_emails)}")


# Iterate over each row in the DataFrame
for index, row in tqdm(df.iterrows(), total=df.shape[0]):
    try:
        # Check if 'Commodity' or 'Location' are NaN or empty strings
        commodity_exists = pd.notna(row["Commodity"]) and row["Commodity"] != ""
        location_exists = pd.notna(row["Location"]) and row["Location"] != ""
        url_contains_http = "http" in row["URL"]

        # Proceed only if both Commodity and Location exist and URL contains 'http'
        if commodity_exists and location_exists and url_contains_http:
            url = row["URL"]
            commodity_value = row["Commodity"]
            location_value = row["Location"]

            # Depending on the URL, call the appropriate function with retries
            if "commodityView" in url:
                with_retry(
                    get_and_save_special_screenshots,
                    url,
                    index,
                    commodity_value,
                    location_value,
                )
            else:
                with_retry(
                    get_and_save_images, url, index, commodity_value, location_value
                )

            # Wait for the specified time before processing the next row
            time.sleep(wait_time_between_rows)
    except Exception as e:
        print(f"An error occurred while processing row {index}: {e}")

# Call the function to merge PDFs in the images/{date} directory
merge_pdfs_in_row_folders(images_base_dir)
