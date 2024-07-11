# KRMS Device Filter and Email Report

This project filters device data from the KRMS API, enriches it with geolocation information, and sends a summary report via email.

## Features

- Fetches device data from the KRMS API.
- Enriches device data with geolocation information using IP addresses.
- Generates a summary report of filtered devices.
- Saves the filtered data to CSV and Excel files.
- Sends an email with the summary report and Excel file attachment.

## Requirements

- Python 3.7+
- `requests` library
- `pandas` library
- `python-dotenv` library
- `openpyxl` library
- `xlsxwriter` library
- `smtplib` library (standard in Python)

## Installation

1. Clone the repository:

    ```sh
    git clone https://github.com/Marzlio/KRMS.git
    cd KRMS
    ```

2. Create a virtual environment and activate it:

    ```sh
    python -m venv venv
    source venv/bin/activate # On Windows use `venv\Scripts\activate`
    ```

3. Install the required packages:

    ```sh
    pip install -r requirements.txt
    ```

4. Create a `.env` file in the project root with the following content:

    ```env
    API_USERNAME=your_api_username
    PASSWORD=your_api_password
    CLIENT_KEY=your_client_key
    PAGE=1
    LIMIT=500000
    ORDERS=[]
    IPINFO_TOKEN=your_ipinfo_token
    XLSX_OUTPUT_FILE=KRMS_Devices.xlsx
    CSV_OUTPUT_FILE=KRMS_Devices.csv
    SMTP_SERVER=smtp.office365.com
    SMTP_PORT=587
    TTLS=TRUE
    LOGIN_REQUIRED=TRUE
    EMAIL_USERNAME=your_office365_email@example.com
    EMAIL_PASSWORD=your_office365_password
    EMAIL_TO=recipient1@example.com,recipient2@example.com
    EMAIL_SUBJECT=KRMS Devices Report
    SEND_EMAIL=true
    ```

## Usage

1. Run the script:

    ```sh
    python KRMS_getdata.py
    ```

    This will fetch the device data, enrich the data with geolocation information, save the filtered data to CSV and Excel files, and send an email with the report if `SEND_EMAIL` is set to `true`.

## File Structure

- `KRMS_getdata.py`: Main script to fetch, filter, enrich data, and send email.
- `.env`: Environment variables for configuration (not included in the repo).
- `requirements.txt`: List of required Python packages.
- `README.md`: This file.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
