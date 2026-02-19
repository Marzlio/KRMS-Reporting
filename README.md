# KRMS Device Filter and Email Report

This project filters device data from the KRMS API, enriches it with geolocation information, and sends a summary report via email.

## Features

- Fetches device data from the KRMS API.
- Enriches device data with geolocation information using IP addresses.
- Generates a summary report of filtered devices.
- Saves the filtered data to CSV and Excel files (with timestamp in filenames).
- Sends an email with the summary report and Excel file attachment.
- Keeps only the N most recent output runs (default 10) to avoid filling the disk.

## Requirements

- Python 3.7+
- `requests`, `pandas`, `python-dotenv`, `openpyxl`, `xlsxwriter`, `geoip2`

## Installation

1. Clone the repository:

    ```sh
    git clone https://github.com/Marzlio/KRMS-Reporting.git
    cd KRMS-Reporting
    ```

2. Create a virtual environment and activate it:

    ```sh
    python -m venv venv
    source venv/bin/activate   # On Windows: venv\Scripts\activate
    ```

3. Install the required packages:

    ```sh
    pip install -r requirements.txt
    ```

4. Copy `.env.sample` to `.env` and fill in your credentials (API_USERNAME, PASSWORD, CLIENT_KEY, MAXMIND_LICENSE_KEY, and optional email/output settings).

## Usage

```sh
python KRMS_getdata.py
```

This will fetch the device data, enrich with geolocation, write timestamped CSV/XLSX/HTML (e.g. `KRMS_Devices_2026-02-19_053315.xlsx`), clean up outputs older than the 10 most recent runs, and send the report email if `SEND_EMAIL` is true.

## Optional environment variables

- **GEOIP_REFRESH_DAYS** – Comma-separated weekdays (1=Mon … 7=Sun) when to refresh the GeoIP DB; default `2,5` (Tue, Fri).
- **KEEP_LAST_N_RUNS** – Number of most recent timestamped runs to keep; default `10`. Older CSV/XLSX/HTML files are deleted after each run.

## File structure

- `KRMS_getdata.py` – Main script (fetch, filter, enrich, report, email, cleanup).
- `.env` – Environment variables (not in repo); use `.env.sample` as template.
- `requirements.txt` – Python dependencies.
- `README.md` – This file.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
