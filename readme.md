# Crop Report Automation

This repository contains a script for automating the generation of crop reports. The script scrapes data from a Google Sheet and uses the Playwright library to retrieve images or screenshots from web pages. It then generates a PDF report and sends it to specified recipients via email.

[![YouTube Video](https://img.youtube.com/vi/ZMN70wnHQvM/0.jpg)](https://youtu.be/ZMN70wnHQvM)

## Features

- Scrapes data from Google Sheet in CSV format
- Utilizes Playwright to retrieve images or screenshots from web pages
- Generates a PDF report with the collected images or screenshots
- Sends the PDF report via email to specified recipients
- Handles errors and retries failed operations

## Installation

1. Clone the repository:
```bash
git clone https://github.com/your-username/crop-report-automation.git
```

2. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Configuration

Before running the script, make sure to configure the following parameters in the script:

- `smtp_server`: Replace with your SMTP server address
- `port`: Replace with your SMTP server's SSL port number
- `sender_email`: Replace with your email address
- `sender_password`: Replace with your email password
- `recipient_emails`: Replace with recipient email addresses
- `sheet_url`: Replace with the URL of the Google Sheet in CSV format

## Usage

To start the automation process, run the following command:
```bash
python crop_report_automation.py
```

The script will retrieve data from the Google Sheet, scrape images or screenshots from the web pages specified in the sheet, generate PDF reports for each row in the sheet, and merge them into a single PDF report.

The script utilizes the Playwright library to interact with web pages and capture the required images or screenshots. It provides functions to handle image processing, PDF generation, retrying failed operations, and sending emails with attachments.

During the execution of the script, it will create a folder structure based on the rows in the Google Sheet. Each row will have its own folder containing the related images or screenshots. The script will check if the folder already exists, and if so, it will skip the retrieval process for that row.

The generated PDF reports will be saved in the `output` folder. The script will merge all the individual PDF reports into a single PDF report named `crop_report_<current_date>.pdf`. The merged PDF report will be saved as `output/crop_report_<current_date>.pdf`.

Finally, the script will send an email with the generated PDF report attached to the specified recipients using the SMTP server and email credentials configured in the script.

## Dependencies

The script requires the following dependencies:

- playwright
- pandas
- fpdf
- Pillow
- PyPDF2
- tqdm

These dependencies can be installed via `pip` using the `requirements.txt` file provided with the repository.

## Conclusion

The Crop Report Automation script offers a streamlined solution for automating the generation of crop reports. By leveraging Playwright for web scraping and integrating email functionality, it simplifies the process of gathering data, generating reports, and distributing them to relevant parties.
