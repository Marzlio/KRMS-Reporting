import os
import json
import csv
import tarfile
import tempfile
import shutil
import logging
import requests
import pandas as pd
import geoip2.database
import ipaddress
import smtplib

from datetime import datetime, timedelta
from dotenv import load_dotenv
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
from typing import Dict, Any, Tuple, List

# -----------------------------------------------------------------------------
# Load environment variables
# -----------------------------------------------------------------------------
load_dotenv()

# KRMS API credentials
API_USERNAME    = os.getenv('API_USERNAME') or (_ for _ in ()).throw(ValueError("Missing API_USERNAME"))
PASSWORD        = os.getenv('PASSWORD')     or (_ for _ in ()).throw(ValueError("Missing PASSWORD"))
CLIENT_KEY      = os.getenv('CLIENT_KEY')   or (_ for _ in ()).throw(ValueError("Missing CLIENT_KEY"))

# Pagination
PAGE            = int(os.getenv('PAGE', '1'))
LIMIT           = int(os.getenv('LIMIT', '10000000'))
try:
    ORDERS = json.loads(os.getenv('ORDERS', '["syncTime DESC"]'))
except json.JSONDecodeError:
    ORDERS = ["syncTime DESC"]

# Output files
CSV_OUTPUT_FILE  = os.getenv('CSV_OUTPUT_FILE', 'devices.csv')
XLSX_OUTPUT_FILE = os.getenv('XLSX_OUTPUT_FILE', 'devices.xlsx')
REPORT_FILE      = 'krms_devices_report.html'

# GeoIP config
MAXMIND_LICENSE_KEY = os.getenv('MAXMIND_LICENSE_KEY') or (_ for _ in ()).throw(ValueError("Missing MAXMIND_LICENSE_KEY"))
GEOIP_DB_PATH       = os.getenv('GEOIP_DB_PATH', 'GeoLite2-City.mmdb')

# Email settings
SMTP_SERVER     = os.getenv('SMTP_SERVER', '')
SMTP_PORT       = int(os.getenv('SMTP_PORT', '587'))
TTLS            = os.getenv('TTLS', 'true').strip().lower() in ('true','1','yes')
LOGIN_REQUIRED  = os.getenv('LOGIN_REQUIRED', 'true').strip().lower() in ('true','1','yes')
EMAIL_USERNAME  = os.getenv('EMAIL_USERNAME','')
EMAIL_PASSWORD  = os.getenv('EMAIL_PASSWORD','')
EMAIL_TO        = os.getenv('EMAIL_TO','')
EMAIL_SUBJECT   = os.getenv('EMAIL_SUBJECT','KRMS Devices Report')
SEND_EMAIL      = os.getenv('SEND_EMAIL','true').strip().lower() in ('true','1','yes')
ATTACH_FILE     = os.getenv('ATTACH_FILE','true').strip().lower() in ('true','1','yes')

# -----------------------------------------------------------------------------
# Logging setup
# -----------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# -----------------------------------------------------------------------------
# Refresh GeoLite2 database on Tue/Fri or if missing
# -----------------------------------------------------------------------------
def refresh_geolite_db():
    today = datetime.utcnow().date()
    if os.path.exists(GEOIP_DB_PATH) and today.isoweekday() not in (2,5):
        return
    url = (
        f"https://download.maxmind.com/app/geoip_download"
        f"?edition_id=GeoLite2-City"
        f"&license_key={MAXMIND_LICENSE_KEY}"
        f"&suffix=tar.gz"
    )
    logger.info(f"Downloading GeoLite2 database...")
    resp = requests.get(url, stream=True)
    resp.raise_for_status()
    with tempfile.NamedTemporaryFile(delete=False) as tmp:
        for chunk in resp.iter_content(1024*1024):
            tmp.write(chunk)
    with tarfile.open(tmp.name, 'r:gz') as tar:
        member = next(m for m in tar.getmembers() if m.name.endswith('.mmdb'))
        with tar.extractfile(member) as mm, open(GEOIP_DB_PATH, 'wb') as out:
            shutil.copyfileobj(mm, out)
    os.unlink(tmp.name)
    logger.info(f"GeoLite2 updated at {GEOIP_DB_PATH}")

refresh_geolite_db()
geo_reader = geoip2.database.Reader(GEOIP_DB_PATH)

# -----------------------------------------------------------------------------
# KRMS API functions
# -----------------------------------------------------------------------------
def request_token() -> str:
    resp = requests.post(
        "https://www.krms.openview.co.za/auth/v1/token",
        headers={"Content-Type": "application/json;charset=utf-8"},
        json={"user": API_USERNAME, "password": PASSWORD, "clientKey": CLIENT_KEY}
    )
    resp.raise_for_status()
    body = resp.json()
    if body.get('code') != 'success':
        raise RuntimeError("Token request failed")
    logger.info("API token received.")
    return body['token']

# -----------------------------------------------------------------------------
# Fetch all devices
# -----------------------------------------------------------------------------
def fetch_all_data(token: str) -> List[dict]:
    headers = {"Authorization": f"Bearer {token}"}
    # Warmup
    for url in ("https://www.krms.openview.co.za/auth/v1/profile",
                "https://www.krms.openview.co.za/api/v1/iams/user"):  
        r = requests.get(url, headers=headers); r.raise_for_status()
    devices_url = "https://www.krms.openview.co.za/api/v1/devices/connects/page"
    all_devices = []
    page = 1
    while True:
        payload = {"page": page, "limit": LIMIT, "keyword": {}, "orders": ORDERS}
        r = requests.post(devices_url, headers=headers, json=payload)
        r.raise_for_status()
        data = r.json().get('data', [])
        if not data: break
        all_devices.extend(data)
        logger.info(f"Fetched page {page} ({len(data)} records)")
        page += 1
    return all_devices

# -----------------------------------------------------------------------------
# Local GeoIP lookup
# -----------------------------------------------------------------------------
def lookup_geo(ip: str) -> Dict[str, Any]:
    try:
        rec = geo_reader.city(ip)
        # Primary country code or fallback to registered country
        country = rec.country.iso_code or rec.registered_country.iso_code or ''
        return {
            'country': country,
            'province': rec.subdivisions.most_specific.iso_code or '',
            'city': rec.city.name or '',
            'latitude': rec.location.latitude or 0.0,
            'longitude': rec.location.longitude or 0.0
        }
    except Exception:
        return {
            'country': '',
            'province': '',
            'city': '',
            'latitude': 0.0,
            'longitude': 0.0
        }

# -----------------------------------------------------------------------------
# Process devices & export
# -----------------------------------------------------------------------------
def process_devices(devices: List[dict]) -> Tuple[Dict[str, Any], Dict[str, Dict[str,int]]]:
    # Stats init
    stats = dict.fromkeys([
        'total_devices','cas_activated','devices_in_sa','devices_not_in_sa',
        'devices_online','connected_last_24h','new_connected_last_24h',
        'new_connected_last_7_days','new_connected_since_first_of_month'
    ], 0)
    stats['total_devices'] = len(devices)   
    retailers = {}
    now = datetime.utcnow()
    first_mo = now.replace(day=1)

    # Prepare CSV with geo fields
    headers = set().union(*(d.keys() for d in devices))
    headers.update(['country','province','city','latitude','longitude'])
    with open(CSV_OUTPUT_FILE, 'w', newline='', encoding='utf-8') as cf:
        writer = csv.DictWriter(cf, fieldnames=list(headers))
        writer.writeheader()
        for d in devices:
            ip = d.get('locationIp','')
            if ip:
                geo = lookup_geo(ip)
                d.update(geo)
            # CAS
            act = d.get('cpeServiceStatus')
            is_cas = (isinstance(act,bool) and act) or (isinstance(act,str) and act.lower()=='activated')
            if is_cas: stats['cas_activated'] += 1
            # online
            onv = d.get('online')
            is_on = (isinstance(onv,bool) and onv) or (isinstance(onv,str) and onv.lower()=='true')
            if is_on: stats['devices_online'] += 1
            # country
            if d.get('country') == 'ZA': stats['devices_in_sa'] += 1
            # syncTime
            st = d.get('syncTime')
            if st and datetime.fromtimestamp(st) >= now - timedelta(days=1):
                stats['connected_last_24h'] += 1
            # connectedTime
            ct = d.get('connectedTime')
            if ct:
                cdt = datetime.fromtimestamp(ct)
                if cdt >= now - timedelta(days=1): stats['new_connected_last_24h'] += 1
                if cdt >= now - timedelta(days=7): stats['new_connected_last_7_days'] += 1
                if cdt >= first_mo: stats['new_connected_since_first_of_month'] += 1
            # retailer  
            
            r = d.get('retailer') or 'No Retailer Added'
            ret = retailers.setdefault(r, dict(total=0, activated=0,
                                                  cas_in_sa=0, cas_not_in_za=0,
                                                  in_sa=0, online_not_in_za=0))
            ret['total'] += 1
            if is_cas:
                ret['activated'] += 1
                if d.get('country') == 'ZA': ret['cas_in_sa'] += 1
                else: ret['cas_not_in_za'] += 1
            if d.get('country') == 'ZA': ret['in_sa'] += 1
            if d.get('country') not in ('ZA','') and ct: ret['online_not_in_za'] += 1
            writer.writerow(d)

# -----------------------------------------------------------------------------
# Process devices & export
# -----------------------------------------------------------------------------
def process_devices(devices: List[dict]) -> Tuple[Dict[str, Any], Dict[str, Dict[str,int]]]:
    # Stats init
    stats = dict.fromkeys([
        'total_devices','cas_activated','devices_in_sa','devices_not_in_sa',
        'devices_online','connected_last_24h','new_connected_last_24h',
        'new_connected_last_7_days','new_connected_since_first_of_month'
    ], 0)
    stats['total_devices'] = len(devices)
    retailers = {}
    now = datetime.utcnow()
    first_mo = now.replace(day=1)

    # Prepare CSV with geo fields
    headers = set().union(*(d.keys() for d in devices))
    headers.update(['country','province','city','latitude','longitude'])
    with open(CSV_OUTPUT_FILE, 'w', newline='', encoding='utf-8') as cf:
        writer = csv.DictWriter(cf, fieldnames=list(headers))
        writer.writeheader()
        for d in devices:
            ip = d.get('locationIp','')
            if ip:
                geo = lookup_geo(ip)
                d.update(geo)
            # CAS
            act = d.get('cpeServiceStatus')
            is_cas = (isinstance(act,bool) and act) or (isinstance(act,str) and act.lower()=='activated')
            if is_cas: stats['cas_activated'] += 1
            # online
            onv = d.get('online')
            is_on = (isinstance(onv,bool) and onv) or (isinstance(onv,str) and onv.lower()=='true')
            if is_on: stats['devices_online'] += 1
            # country
            if d.get('country') == 'ZA': stats['devices_in_sa'] += 1
            # syncTime
            st = d.get('syncTime')
            if st and datetime.fromtimestamp(st) >= now - timedelta(days=1):
                stats['connected_last_24h'] += 1
            # connectedTime
            ct = d.get('connectedTime')
            if ct:
                cdt = datetime.fromtimestamp(ct)
                if cdt >= now - timedelta(days=1): stats['new_connected_last_24h'] += 1
                if cdt >= now - timedelta(days=7): stats['new_connected_last_7_days'] += 1
                if cdt >= first_mo: stats['new_connected_since_first_of_month'] += 1
            # retailer
            r = d.get('retailer') or 'No Retailer Added'
            ret = retailers.setdefault(r, dict(total=0, activated=0,
                                               cas_in_sa=0, cas_not_in_za=0,
                                               in_sa=0, online_not_in_za=0))
            ret['total'] += 1
            if is_cas:
                ret['activated'] += 1
                if d.get('country') == 'ZA': ret['cas_in_sa'] += 1
                else: ret['cas_not_in_za'] += 1
            if d.get('country') == 'ZA': ret['in_sa'] += 1
            if d.get('country') not in ('ZA','') and ct: ret['online_not_in_za'] += 1
            writer.writerow(d)

    # XLSX export
    pd.DataFrame(devices).to_excel(XLSX_OUTPUT_FILE, index=False)
    stats['devices_not_in_sa'] = stats['cas_activated'] - stats['devices_in_sa']
    return stats, retailers

# -----------------------------------------------------------------------------
# Generate HTML report
# -----------------------------------------------------------------------------
def generate_report(stats: Dict[str,int], retailers: Dict[str,Dict[str,int]]) -> str:
    html = f"""
    <html><head>
      <style>
        body {{ font-family: Arial; margin: 20px; }}
        h1 {{ color: #004080; }}
        h2 {{ margin-top: 30px; color: #004080; }}
        .summary {{ margin-bottom: 20px; }}
        .summary p {{ margin: 4px 0; font-size: 14px; }}
        .summary .positive {{ color: green; font-weight: bold; }}
        .summary .negative {{ color: red;   font-weight: bold; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; font-size: 13px; }}
        th {{ background: #004080; color: #fff; text-align: left; }}
        tbody tr:nth-child(odd) {{ background: #f9f9f9; }}
      </style>
    </head><body>
      <h1>KRMS Devices Report</h1>

      <div class="summary">
        <h2>Summary:</h2>
        <p>Total Number of devices on KRMS: {stats['total_devices']}</p>
        <p>Total Number of CAS activated devices: {stats['cas_activated']}</p>
        <p class="positive">Total Number of devices in South Africa: {stats['devices_in_sa']}</p>
        <p class="negative">Devices not in South Africa: {stats['devices_not_in_sa']}</p>
        <p>Number of devices currently online: {stats['devices_online']}</p>
        <p>Number of devices connected in the last 24 hours: {stats['connected_last_24h']}</p>
        <p>New devices connected in the last 24 hours: {stats['new_connected_last_24h']}</p>
        <p>New devices connected in the last 7 days: {stats['new_connected_last_7_days']}</p>
        <p>New devices connected since the first of the month: {stats['new_connected_since_first_of_month']}</p>
      </div>

      <h2>Devices per retailer:</h2>
      <table>
        <thead>
          <tr>
            <th>Retailer</th>
            <th>Total Devices</th>
            <th>CAS Activated</th>
            <th>CAS Activated not in ZA</th>
            <th>CAS Activated in ZA</th>
            <th>Online in ZA</th>
            <th>Online Not in ZA</th>
          </tr>
        </thead>
        <tbody>
    """
    for retailer, vals in retailers.items():
        html += (
          f"<tr>"
          f"<td>{retailer}</td>"
          f"<td>{vals['total']}</td>"
          f"<td>{vals['activated']}</td>"
          f"<td>{vals['cas_not_in_za']}</td>"
          f"<td>{vals['cas_in_sa']}</td>"
          f"<td>{vals['in_sa']}</td>"
          f"<td>{vals['online_not_in_za']}</td>"
          f"</tr>"
        )
    html += """
        </tbody>
      </table>
    </body></html>
    """
    with open(REPORT_FILE, 'w', encoding='utf-8') as f:
        f.write(html)
    return html

# -----------------------------------------------------------------------------
# Send email
# -----------------------------------------------------------------------------
def send_email(report: str, attachment: str) -> None:
    if not (EMAIL_USERNAME and EMAIL_PASSWORD and SMTP_SERVER and EMAIL_TO):
        logger.error("Email not configured; skipping send.")
        return
    msg = MIMEMultipart()
    msg['From'] = EMAIL_USERNAME
    msg['To'] = EMAIL_TO
    msg['Subject'] = EMAIL_SUBJECT
    msg.attach(MIMEText(report, 'html'))
    if ATTACH_FILE and os.path.exists(attachment):
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(open(attachment, 'rb').read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment)}')
        msg.attach(part)
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        if TTLS: server.starttls()
        if LOGIN_REQUIRED: server.login(EMAIL_USERNAME, EMAIL_PASSWORD)
        server.sendmail(EMAIL_USERNAME, EMAIL_TO.split(','), msg.as_string())
    logger.info("Email sent.")

# -----------------------------------------------------------------------------
# Main execution
# -----------------------------------------------------------------------------
if __name__ == '__main__':
    logger.info("Script starting.")
    token = request_token()
    devices = fetch_all_data(token)
    if not devices:
        logger.info("No devices found.")
        exit(0)
    stats, retailers = process_devices(devices)
    report_html = generate_report(stats, retailers)
    if SEND_EMAIL:
        send_email(report_html, XLSX_OUTPUT_FILE)
    logger.info("Script completed.")