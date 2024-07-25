# -*- coding: utf-8 -*-
import requests
import csv
import pandas as pd
from datetime import datetime, timedelta
from dotenv import load_dotenv
import os
import json
import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from typing import Dict, Any

# Load environment variables from .env file
load_dotenv()

# Constants
API_USERNAME = os.getenv('API_USERNAME')
PASSWORD = os.getenv('PASSWORD')
CLIENT_KEY = os.getenv('CLIENT_KEY')
PAGE = int(os.getenv('PAGE', 1))
LIMIT = int(os.getenv('LIMIT', 10000000))
ORDERS = json.loads(os.getenv('ORDERS', '["syncTime DESC"]'))
CSV_OUTPUT_FILE = os.getenv('CSV_OUTPUT_FILE', 'devices.csv')
XLSX_OUTPUT_FILE = os.getenv('XLSX_OUTPUT_FILE', 'devices.xlsx')
IPINFO_TOKEN = os.getenv('IPINFO_TOKEN')
SMTP_SERVER = os.getenv('SMTP_SERVER')
SMTP_PORT = int(os.getenv('SMTP_PORT', 587))
TTLS = os.getenv('TTLS', 'TRUE').upper() == 'TRUE'
LOGIN_REQUIRED = os.getenv('LOGIN_REQUIRED', 'TRUE').upper() == 'TRUE'
EMAIL_USERNAME = os.getenv('EMAIL_USERNAME')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
EMAIL_TO = os.getenv('EMAIL_TO').split(',')
EMAIL_SUBJECT = os.getenv('EMAIL_SUBJECT')
SEND_EMAIL = os.getenv('SEND_EMAIL', 'TRUE').upper() == 'TRUE'

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
logger = logging.getLogger()

# File for the final report
report_file = 'krms_devices_report.txt'

class IPFetchError(Exception):
    """Custom exception for IP fetch failures."""
    pass

def request_token() -> str:
    """Request an API token."""
    token_url = "https://www.krms.openview.co.za/auth/v1/token"
    headers = {
        "Content-Type": "application/json;charset=utf-8",
        "User-Agent": "Mozilla/5.0"
    }
    data = {
        "user": API_USERNAME,
        "password": PASSWORD,
        "clientKey": CLIENT_KEY
    }
    try:
        response = requests.post(token_url, headers=headers, json=data)
        response.raise_for_status()  # Raise an exception for HTTP errors
        token_response = response.json()
        if token_response.get("code") == "success":
            logging.info("Token received successfully.")
            return token_response.get("token")
        else:
            logging.error("Failed to get token. Response code: %s", token_response.get("code"))
            raise Exception("Failed to get token.")
    except requests.RequestException as e:
        logging.error("Error requesting token: %s", e)
        raise

def request_data(url: str, headers: dict, data: dict = None) -> dict:
    """Request data from the API."""
    try:
        response = requests.post(url, headers=headers, json=data) if data else requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        logging.error("Error requesting data from %s: %s", url, e)
        raise

def load_ip_info(file_path: str = 'ip_info.json') -> Dict[str, Any]:
    """Load IP info from a JSON file."""
    try:
        if os.path.exists(file_path):
            with open(file_path, 'r') as file:
                logging.info("Loaded IP info from %s", file_path)
                return json.load(file)
    except Exception as e:
        logging.error("Error loading IP info: %s", e)
    return {}

def save_ip_info(ip_info: Dict[str, Any], file_path: str = 'ip_info.json') -> None:
    """Save IP info to a JSON file."""
    try:
        with open(file_path, 'w') as file:
            json.dump(ip_info, file)
            logging.info("Saved IP info to %s", file_path)
    except Exception as e:
        logging.error("Error saving IP info: %s", e)

def fetch_ip_info(ip_address: str, ip_info_cache: Dict[str, Any]) -> Dict[str, Any]:
    """Fetch IP information, using cache if available."""
    if ip_address in ip_info_cache:
        return ip_info_cache[ip_address]

    url = f"https://ipinfo.io/{ip_address}?token={IPINFO_TOKEN}"
    try:
        response = requests.get(url)
        response.raise_for_status()
        ip_info_cache[ip_address] = response.json()
        save_ip_info(ip_info_cache)
        return ip_info_cache[ip_address]
    except requests.RequestException as e:
        raise IPFetchError(f"Failed to fetch data for {ip_address}: {e}")  # Raise custom exception

def generate_report(stats: Dict[str, Any], retailers: Dict[str, Dict[str, int]]) -> str:
    """Generate a report with the collected statistics and return the report as a string."""
    report_content = f"""
    <html>
        <head>
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    color: #333;
                }}
                h1 {{
                    color: #004080;
                }}
                p {{
                    margin: 0 0 10px;
                }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                }}
                th, td {{
                    border: 1px solid #ddd;
                    padding: 8px;
                }}
                th {{
                    padding-top: 12px;
                    padding-bottom: 12px;
                    text-align: left;
                    background-color: #004080;
                    color: white;
                }}
                .section {{
                    margin-bottom: 20px;
                }}
                .section-title {{
                    font-weight: bold;
                    text-decoration: underline;
                }}
            </style>
        </head>
        <body>
            <h1>KRMS Devices Report</h1>
            <div class="section">
                <p class="section-title">Summary:</p>
                <p>Total Number of devices on KRMS: {stats['total_devices']}</p>
                <p>Total Number of CAS activated devices: {stats['cas_activated']}</p>
                <p>Total Number of devices in South Africa: <span style="color: green; font-weight: bold;">{stats['devices_in_sa']}</span></p>
                <p>Devices not in South Africa: <span style="color: red; font-weight: bold;">{stats['devices_not_in_sa']}</span></p>
                <p>Number of devices currently online: {stats['devices_online']}</p>
                <p>Number of devices connected in the last 24 hours: {stats['connected_last_24h']}</p>
                <p>New devices connected in the last 24 hours: {stats['new_connected_last_24h']}</p>
                <p>New devices connected in the last 7 days: {stats['new_connected_last_7_days']}</p>
                <p>New devices connected since the first of the month: {stats['new_connected_since_first_of_month']}</p>
            </div>
            <div class="section">
                <p class="section-title">Devices per retailer:</p>
                <table>
                    <thead>
                        <tr>
                            <th>Retailer</th>
                            <th>CAS Activated</th>
                            <th>Total Devices</th>
                            <th>Online in South Africa</th>
                        </tr>
                    </thead>
                    <tbody>
    """
    for retailer, counts in retailers.items():
        report_content += f"""
                        <tr>
                            <td>{retailer}</td>
                            <td>{counts['activated']}</td>
                            <td>{counts['total']}</td>
                            <td>{counts['in_sa']}</td>
                        </tr>
        """
    
    report_content += """
                    </tbody>
                </table>
            </div>
        </body>
    </html>
    """
    
    with open(report_file, 'w') as report:
        report.write(report_content)

    return report_content

def send_email(report_content: str, attachment_path: str) -> None:
    """Send an email with the report and attachment."""
    msg = MIMEMultipart()
    msg['From'] = EMAIL_USERNAME
    msg['To'] = ', '.join(EMAIL_TO)
    msg['Subject'] = EMAIL_SUBJECT

    # Attach the body with the msg instance
    msg.attach(MIMEText(report_content, 'html'))

    # Open the file to be sent
    with open(attachment_path, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())

    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(attachment_path)}")
    msg.attach(part)

    # Send the message via the SMTP server
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        if TTLS:
            server.starttls()
        if LOGIN_REQUIRED:
            server.login(EMAIL_USERNAME, EMAIL_PASSWORD)
        server.sendmail(EMAIL_USERNAME, EMAIL_TO, msg.as_string())
        logging.info("Email sent successfully.")
        server.quit()
    except Exception as e:
        logging.error(f"Failed to send email: {e}")

def main() -> None:
    """Main function to execute the script."""
    logging.info("Script started.")
    try:
        token = request_token()
    except Exception as e:
        logging.error("Failed to obtain token: %s", e)
        return

    profile_url = "https://www.krms.openview.co.za/auth/v1/profile"
    user_url = "https://www.krms.openview.co.za/api/v1/iams/user"
    devices_url = "https://www.krms.openview.co.za/api/v1/devices/connects/page"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json;charset=utf-8",
        "User-Agent": "Mozilla/5.0"
    }

    try:
        logging.info("Requesting profile data...")
        profile_data = request_data(profile_url, headers)
        logging.info("Profile data received.")
        
        logging.info("Requesting user data...")
        user_data = request_data(user_url, headers)
        logging.info("User data received.")
        
        logging.info("Requesting devices data...")
        devices_data = request_data(devices_url, headers, data={
            "page": PAGE,
            "limit": LIMIT,
            "keyword": {},
            "orders": ORDERS,
        })
        logging.info("Devices data received.")
    except requests.RequestException as e:
        logging.error("Error requesting data: %s", e)
        return

    # Load IP info cache
    ip_info_cache = load_ip_info()
    devices = devices_data.get('data', [])

    if not devices:
        logging.info("No devices data found.")
        return

    csv_headers = list(devices[0].keys()) if devices else []
    if "country" not in csv_headers:
        csv_headers.append("country")

    updated_devices = []

    # Initialize statistics
    stats = {
        'total_devices': len(devices),
        'cas_activated': 0,
        'devices_in_sa': 0,
        'devices_not_in_sa': 0,
        'devices_online': 0,
        'connected_last_24h': 0,
        'new_connected_last_24h': 0,
        'new_connected_last_7_days': 0,
        'new_connected_since_first_of_month': 0
    }
    retailers = {}

    now = datetime.now()
    first_of_month = now.replace(day=1)

    with open(CSV_OUTPUT_FILE, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=csv_headers)
        writer.writeheader()

        for device in devices:
            device_id = device.get('device_id')  # Use 'device_id'

            if not device_id:
                logging.warning(f"Skipping device with missing ID. Full data: {device}")
                continue

            ip_address = device.get('locationIp')
            if ip_address:
                try:
                    ip_info = fetch_ip_info(ip_address, ip_info_cache)  # Handle IPFetchError
                except IPFetchError as e:
                    logging.error(str(e))
                    continue  # Skip to the next device if IP fetch fails

                if 'error' not in ip_info:
                    ip_province = ip_info.get('region')
                    ip_city = ip_info.get('city')
                    ip_latitude, ip_longitude = map(float, ip_info.get('loc', '0,0').split(','))
                    ip_country = ip_info.get('country')

                    # Compare and update logic
                    if (device.get('province') != ip_province or 
                        device.get('city') != ip_city or 
                        device.get('latitude') != ip_latitude or 
                        device.get('longitude') != ip_longitude or
                        device.get('country') != ip_country):
                        
                        device['province'] = ip_province
                        device['city'] = ip_city
                        device['latitude'] = ip_latitude
                        device['longitude'] = ip_longitude
                        device['country'] = ip_country

            # Update statistics
            activation_status = device.get('cpeServiceStatus')
            if (isinstance(activation_status, bool) and activation_status) or (isinstance(activation_status, str) and activation_status.lower() == 'activated'):
                stats['cas_activated'] += 1

            if device.get('country') == 'ZA':
                stats['devices_in_sa'] += 1

            if (isinstance(device.get('online'), bool) and device['online']) or (isinstance(device.get('online'), str) and device['online'].lower() == 'true'):
                stats['devices_online'] += 1

            # Check connection times
            sync_time = device.get('syncTime')
            connected_time = device.get('connectedTime')

            if sync_time:
                sync_time = datetime.fromtimestamp(sync_time)
                if sync_time >= now - timedelta(days=1):
                    stats['connected_last_24h'] += 1

            if connected_time:
                connected_time = datetime.fromtimestamp(connected_time)
                if connected_time >= now - timedelta(days=1):
                    stats['new_connected_last_24h'] += 1
                if connected_time >= now - timedelta(days=7):
                    stats['new_connected_last_7_days'] += 1
                if connected_time >= first_of_month:
                    stats['new_connected_since_first_of_month'] += 1

            retailer = device.get('retailer', 'No Retailer Added')
            if retailer in retailers:
                retailers[retailer]['total'] += 1
                if (isinstance(activation_status, bool) and activation_status) or (isinstance(activation_status, str) and activation_status.lower() == 'activated'):
                    retailers[retailer]['activated'] += 1
                if device.get('country') == 'ZA':
                    retailers[retailer]['in_sa'] += 1
            else:
                retailers[retailer] = {
                    'total': 1,
                    'activated': 1 if (isinstance(activation_status, bool) and activation_status) or (isinstance(activation_status, str) and activation_status.lower() == 'activated') else 0,
                    'in_sa': 1 if device.get('country') == 'ZA' else 0
                }

            writer.writerow(device)

    # Save data to XLSX
    df = pd.DataFrame(devices)
    df.to_excel(XLSX_OUTPUT_FILE, index=False)

    stats['devices_not_in_sa'] = stats['cas_activated'] - stats['devices_in_sa']

    logging.info(f"Data successfully exported to {CSV_OUTPUT_FILE} and {XLSX_OUTPUT_FILE}")
    logging.info("Script completed.")

    # Generate and log report
    report_content = generate_report(stats, retailers)

    # Send email if required
    if SEND_EMAIL:
        send_email(report_content, XLSX_OUTPUT_FILE)

if __name__ == "__main__":
    main()