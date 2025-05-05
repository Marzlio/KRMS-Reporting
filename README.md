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
- `geoip2` library

## Installation

1. Clone the repository:

    ```sh
    git clone https://github.com/Marzlio/KRMS-Reporting.git
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
# KRMS API
API_USERNAME=<your KRMS username>
PASSWORD=<your KRMS password>
CLIENT_KEY=<your KRMS client key>

# Pagination (optional)
PAGE=1
LIMIT=10000000
ORDERS=[]

# Output file names (optional)
CSV_OUTPUT_FILE=devices.csv
XLSX_OUTPUT_FILE=devices.xlsx

# GeoIP
MAXMIND_LICENSE_KEY=<your MaxMind license-key>
GEOIP_DB_PATH=GeoLite2-City.mmdb

# Email (if you want the script to send mail)
SMTP_SERVER=<smtp.example.com>
SMTP_PORT=587
TTLS=true              # start TLS?
LOGIN_REQUIRED=true
EMAIL_USERNAME=<smtp login>
EMAIL_PASSWORD=<smtp password>
EMAIL_TO="user@domain.com,other@domain.com"
EMAIL_SUBJECT="KRMS Devices Report"
SEND_EMAIL=true
ATTACH_FILE=true
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
