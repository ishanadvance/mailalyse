# Email Processing

This project is a Python script that connects to an IMAP server, retrieves emails, and saves their data into an Excel file. It can be used to automate email processing tasks, such as extracting sender information, subject lines, dates, and attachments.

## Prerequisites

Before running the script, make sure you have the following:

- Python 3.x installed on your machine
- Required packages: `imaplib`, `email`, `datetime`, `openpyxl`

## Configuration

To use the script, you need to provide the following configuration:

- **IMAP Configuration:**
  - `IMAP_SERVER`: The IMAP server address (e.g., imap.gmail.com)
  - `IMAP_PORT`: The IMAP server port (e.g., 993 for SSL/TLS)

- **Email Account Configuration:**
  - `EMAIL_USERNAME`: Your email address
  - `EMAIL_PASSWORD`: Your email account password

- **Excel Configuration:**
  - `EXCEL_FILENAME`: The name of the Excel file to store email data
  - `EXCEL_SHEETNAME`: The name of the Excel sheet within the file

- **Progress Tracking Configuration:**
  - `BATCH_SIZE`: The number of emails to process in each batch

Make sure to update these configuration values in the script according to your requirements before running it.

## Running the Script

To run the script, follow these steps:

1. Clone the repository or download the `email_processing.py` file to your local machine.
2. Install the required packages using pip: `pip install imaplib email openpyxl`
3. Open the `email_processing.py` file in a text editor.
4. Update the configuration values according to your email and Excel settings.
5. Save the changes to the file.
6. Open a terminal or command prompt and navigate to the directory where the script is located.
7. Run the script using the following command: `python email_processing.py` or `python3 email_processing.py`

The script will connect to the specified IMAP server, retrieve the emails, process them, and save the extracted data into an Excel file.

## Output

The script saves the extracted email data into an Excel file with the specified name (`EXCEL_FILENAME`). The data is stored in the Excel sheet (`EXCEL_SHEETNAME`) in a tabular format, including the email ID, sender, subject, date, attachment names, and attachment sizes.

## Notes

- Ensure that you have provided correct IMAP server details and valid email account credentials.
- The script uses SSL/TLS to connect to the IMAP server (`IMAP_PORT = 993`). Modify it accordingly if your server uses a different port or protocol.
- If any errors occur during the email processing, the script will log the error and continue with the remaining emails.
- Remember to save the Excel file after the script completes processing all the emails.
