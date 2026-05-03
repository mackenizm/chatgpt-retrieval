import imaplib
import email
from email.policy import default
from email.header import decode_header
from datetime import datetime
import os
import re
from dotenv import load_dotenv
import csv
import base64
import hashlib
import json
import time
from urllib.parse import quote
from bs4 import BeautifulSoup
import fitz  # PyMuPDF
import msal
import requests
from openai import OpenAI

# Load environment variables
load_dotenv()

def env_value(name, default=None):
    value = os.getenv(name, default)
    return value.strip() if isinstance(value, str) else value

# Load API key and email credentials from .env file
OPENAI_API_KEY = env_value('OPENAI_API_KEY')
EMAIL = env_value('EMAIL')
APP_PASSWORD = env_value('APP_PASSWORD')
IMAP_HOST = env_value('IMAP_HOST', 'imap-mail.outlook.com')
IMAP_PORT = int(env_value('IMAP_PORT', '993'))
EMAIL_AUTH_MODE = env_value('EMAIL_AUTH_MODE', 'graph').lower()
MS_CLIENT_ID = env_value('MS_CLIENT_ID')
MS_TENANT = env_value('MS_TENANT', 'consumers')
GRAPH_SCOPES = ['Mail.Read']
GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0'
TOKEN_CACHE_PATH = './data/msal_token_cache.json'

# Define file paths
email_path = './data/email_content.txt'
csv_path = './data/invoice.csv'
dedup_csv_path = './data/dedup_invoice.csv'
pdf_path = './data/pdf_content.txt'

SUBJECT_KEYWORDS = [
    'invoice',
    'receipt',
    'payment',
    'paid',
    'purchase',
    'order confirmation',
    'order confirmed',
    'tax invoice',
    'tax statement',
    'booking confirmed',
    'confirmation',
]

# Define email search date range
# IMAP BEFORE is exclusive, so the end date must be the day after the FY ends.
start_date = env_value("FY_START_DATE", "01-Jul-2024")
end_date = env_value("FY_END_BEFORE_DATE", "01-Jul-2025")

CSV_FIELDS = [
    'Date',
    'Category',
    'Supplier',
    'Invoice Number',
    'Item/Product/Service',
    'Price',
    'GST',
    'Total',
    'Possible Duplicate',
    'Duplicate Group',
    'Duplicate Reason',
    'Possible Duplicate Of',
    'Source Subject',
    'Source Message ID',
    'Source Date',
    'Source Attachments',
]

# Define the debug mode (can be controlled via an environment variable)
DEBUG_MODE = env_value('DEBUG_MODE', 'False').lower() in ('true', '1', 't')
if DEBUG_MODE:
    print(f"DEBUG_MODE is set to: {DEBUG_MODE}")

def debug_print(message):
    if DEBUG_MODE:
        print(message)

def initialize_output_files():
    os.makedirs(os.path.dirname(email_path), exist_ok=True)
    for file_path in (email_path, csv_path, dedup_csv_path, pdf_path):
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write("")

def decode_mime_words(s):
    if not s:
        return ''
    decoded_fragments = email.header.decode_header(s)
    result = ''
    for fragment, encoding in decoded_fragments:
        if isinstance(fragment, bytes):
            try:
                result += fragment.decode(encoding or 'utf-8', errors='replace')
            except LookupError:
                result += fragment.decode('utf-8', errors='replace')
        else:
            result += fragment
    return result

def parse_config_date(value):
    for date_format in ('%d-%b-%Y', '%Y-%m-%d', '%d/%m/%Y'):
        try:
            return datetime.strptime(value, date_format)
        except ValueError:
            continue
    raise ValueError(f"Could not parse date '{value}'. Use formats like 01-Jul-2024 or 2024-07-01.")

def graph_datetime(value):
    return parse_config_date(value).strftime('%Y-%m-%dT00:00:00Z')

def load_token_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_PATH):
        with open(TOKEN_CACHE_PATH, 'r', encoding='utf-8') as file:
            cache.deserialize(file.read())
    return cache

def save_token_cache(cache):
    if cache.has_state_changed:
        os.makedirs(os.path.dirname(TOKEN_CACHE_PATH), exist_ok=True)
        with open(TOKEN_CACHE_PATH, 'w', encoding='utf-8') as file:
            file.write(cache.serialize())

def get_graph_access_token():
    if not MS_CLIENT_ID:
        raise RuntimeError(
            "MS_CLIENT_ID must be set in .env for Microsoft Graph OAuth. "
            "Create a public/native Microsoft app registration, add delegated Mail.Read, "
            "then put its Application (client) ID in MS_CLIENT_ID."
        )

    authority = f"https://login.microsoftonline.com/{MS_TENANT}"
    cache = load_token_cache()
    app = msal.PublicClientApplication(
        MS_CLIENT_ID,
        authority=authority,
        token_cache=cache,
    )

    result = None
    accounts = app.get_accounts(username=EMAIL) if EMAIL else app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
        if 'user_code' not in flow:
            raise RuntimeError(f"Could not start Microsoft device login: {flow}")
        print(flow['message'])
        result = app.acquire_token_by_device_flow(flow)

    save_token_cache(cache)
    if 'access_token' not in result:
        error = result.get('error_description') or result.get('error') or result
        raise RuntimeError(f"Microsoft Graph OAuth failed: {error}")
    return result['access_token']

def graph_get(url, token, params=None):
    response = requests.get(
        url,
        headers={'Authorization': f'Bearer {token}'},
        params=params,
        timeout=60,
    )
    if response.status_code >= 400:
        raise RuntimeError(f"Microsoft Graph request failed ({response.status_code}): {response.text}")
    return response.json()

def fetch_graph_attachments(message_id, token):
    attachments = []
    url = f"{GRAPH_BASE_URL}/me/messages/{quote(message_id, safe='')}/attachments"
    params = None

    while url:
        data = graph_get(url, token, params=params)
        for attachment in data.get('value', []):
            filename = attachment.get('name') or ''
            content_bytes = attachment.get('contentBytes')
            if (
                attachment.get('@odata.type') == '#microsoft.graph.fileAttachment'
                and not attachment.get('isInline')
                and filename.lower().endswith('.pdf')
                and content_bytes
            ):
                attachments.append({
                    'filename': filename,
                    'content': base64.b64decode(content_bytes),
                })
        url = data.get('@odata.nextLink')

    return attachments

def fetch_emails_graph():
    token = get_graph_access_token()
    start_iso = graph_datetime(start_date)
    end_iso = graph_datetime(end_date)
    date_filter = f"receivedDateTime ge {start_iso} and receivedDateTime lt {end_iso}"
    params = {
        '$select': 'id,subject,body,receivedDateTime,internetMessageId,hasAttachments',
        '$filter': date_filter,
        '$orderby': 'receivedDateTime asc',
        '$top': '50',
    }
    url = f"{GRAPH_BASE_URL}/me/mailFolders/inbox/messages"

    messages = []
    while url:
        data = graph_get(url, token, params=params)
        params = None
        for item in data.get('value', []):
            subject = item.get('subject') or ''
            if not any(keyword in subject.lower() for keyword in SUBJECT_KEYWORDS):
                continue
            attachments = fetch_graph_attachments(item['id'], token) if item.get('hasAttachments') else []
            messages.append({
                'source': 'graph',
                'subject': subject,
                'body': (item.get('body') or {}).get('content') or '',
                'body_type': (item.get('body') or {}).get('contentType') or '',
                'message_id': item.get('internetMessageId') or item.get('id') or '',
                'date': item.get('receivedDateTime') or '',
                'attachments': attachments,
            })
        url = data.get('@odata.nextLink')

    debug_print(f"Total Graph emails fetched: {len(messages)}")
    return messages

def fetch_emails_imap():
    if not EMAIL or not APP_PASSWORD:
        raise RuntimeError("EMAIL and APP_PASSWORD must be set in .env before fetching emails.")

    # Adding a delay before login attempt
    time.sleep(10)  # Delay for 10 seconds

    mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    try:
        mail.login(EMAIL, APP_PASSWORD)
    except imaplib.IMAP4.error as e:
        raise RuntimeError(
            "IMAP login failed. Check that EMAIL is the full mailbox address, "
            "APP_PASSWORD is an app password or valid IMAP password, IMAP is enabled "
            "for the mailbox, and IMAP_HOST/IMAP_PORT match your email provider."
        ) from e
    mail.select("inbox")

    # Use the date variables in the search criteria
    search_criteria = f'(SINCE "{start_date}" BEFORE "{end_date}")'
    debug_print(f"IMAP search criteria: {search_criteria}")
    result, data = mail.search(None, search_criteria)

    email_ids = data[0].split()
    debug_print(f"Email IDs fetched: {email_ids}")

    emails = []
    for e_id in email_ids:
        result, msg_data = mail.fetch(e_id, "(RFC822)")
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)
        emails.append(msg)
    mail.logout()
    debug_print(f"Total emails fetched: {len(emails)}")
    return emails

def fetch_emails():
    if EMAIL_AUTH_MODE == 'imap':
        return fetch_emails_imap()
    return fetch_emails_graph()

def get_body(msg):
    debug_print("Extracting email body...")
    if isinstance(msg, dict) and msg.get('source') == 'graph':
        return msg.get('body') or ''

    plain_body = None
    html_body = None

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))

            if "attachment" in content_disposition:
                continue

            if content_type == "text/plain":
                try:
                    plain_body = part.get_payload(decode=True).decode('utf-8', errors='replace')
                except UnicodeDecodeError:
                    try:
                        plain_body = part.get_payload(decode=True).decode('latin-1', errors='replace')
                    except UnicodeDecodeError:
                        plain_body = None
            elif content_type == "text/html":
                try:
                    html_body = part.get_payload(decode=True).decode('utf-8', errors='replace')
                except UnicodeDecodeError:
                    try:
                        html_body = part.get_payload(decode=True).decode('latin-1', errors='replace')
                    except UnicodeDecodeError:
                        html_body = None
        return plain_body or html_body or ''
    else:
        try:
            return msg.get_payload(decode=True).decode('utf-8', errors='replace')
        except UnicodeDecodeError:
            try:
                return msg.get_payload(decode=True).decode('latin-1', errors='replace')
            except UnicodeDecodeError:
                return ''

def get_attachments(msg):
    debug_print("Extracting email attachments...")
    if isinstance(msg, dict) and msg.get('source') == 'graph':
        attachments = msg.get('attachments') or []
        debug_print(f"Attachments extracted: {len(attachments)}")
        return attachments

    attachments = []
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        filename = decode_mime_words(part.get_filename())
        if filename and filename.lower().endswith(".pdf"):
            attachments.append({
                'filename': filename,
                'content': part.get_payload(decode=True)
            })
    debug_print(f"Attachments extracted: {len(attachments)}")
    return attachments

def save_pdf_text_to_file(text, file_path):
    debug_print("Saving pdf text to file...")
    with open(file_path, 'a', encoding='utf-8') as file:
        file.write(text + "\n\n---\n\n")

def extract_text_from_pdf(pdf_bytes, filename):
    debug_print(f"Extracting text from PDF {filename}...")
    text = ""
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as pdf:
            for page_num in range(len(pdf)):
                page = pdf.load_page(page_num)
                text += page.get_text()
        save_pdf_text_to_file(text, pdf_path)
    except Exception as e:
        debug_print(f"Error extracting text from PDF {filename}: {e}")
    return text

def clean_html(html_content):
    debug_print("Cleaning HTML content...")
    soup = BeautifulSoup(html_content, 'html.parser')

    # Remove all style and script tags
    for tag in soup(['style', 'script']):
        tag.decompose()

    # Extract the text content
    text = soup.get_text(separator='\n')

    # Remove excessive newlines and whitespace
    cleaned_text = "\n".join(line.strip() for line in text.splitlines() if line.strip())

    return cleaned_text

def save_combined_content_to_file(combined_content, file_path):
    debug_print("Saving combined content to file...")
    with open(file_path, 'a', encoding='utf-8') as file:
        file.write(combined_content + "\n\n---\n\n")

def contains_dollar_amount_and_date(text):
    # Check for dollar amount
    amount_pattern = r'(?:A\$|\$|AUD\s*)\s?\d{1,3}(?:,\d{3})*(?:\.\d{2})?|\d{1,3}(?:,\d{3})+\.\d{2}'
    has_dollar_amount = bool(re.search(amount_pattern, text, re.IGNORECASE))
    # Check for various date formats
    date_patterns = [
        r'\d{4}-\d{2}-\d{2}',  # YYYY-MM-DD
        r'\d{2}/\d{2}/\d{4}',  # DD/MM/YYYY
        r'\d{2}-\d{2}-\d{4}',  # DD-MM-YYYY
        r'\d{2}\.\d{2}\.\d{4}',  # DD.MM.YYYY
        r'\d{1,2} \w+ \d{4}',  # 8 Jan 2024 or 08 Jan 2024
        r'\w+ \d{1,2}, \d{4}', # Jan 8, 2024 or January 8, 2024
    ]
    has_date = any(bool(re.search(pattern, text)) for pattern in date_patterns)
    return has_dollar_amount and has_date

def strip_json_code_fence(text):
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r'^```(?:json)?\s*', '', text, flags=re.IGNORECASE)
        text = re.sub(r'\s*```$', '', text)
    return text.strip()

def parse_json_array(text):
    text = strip_json_code_fence(text)
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r'\[[\s\S]*\]', text)
        if not match:
            raise
        parsed = json.loads(match.group(0))

    if isinstance(parsed, dict):
        for value in parsed.values():
            if isinstance(value, list):
                parsed = value
                break

    if not isinstance(parsed, list):
        raise ValueError("Parsed JSON is not a list")
    return parsed

def extract_details_from_content(subject, body, attachments):
    if re.search(r'<(?:html|body|div|span|p|br)\b', body or '', re.IGNORECASE):
        cleaned_body = clean_html(body)
    else:
        cleaned_body = body
    
    combined_content = f"Subject: {subject}\nBody: {cleaned_body}\n"
        
    for attachment in attachments:
        filename, content = attachment['filename'], attachment['content']
        if filename.lower().endswith('.pdf'):
            pdf_text = extract_text_from_pdf(content, filename)
            combined_content += f"\nAttachment ({filename}):\n{pdf_text}\n"
        else:
            combined_content += f"\nAttachment ({filename}):\n{content}\n"
    
    # Save combined content to a text file
    save_combined_content_to_file(combined_content, email_path)

    # Check for presence of dollar amount and date
    if not contains_dollar_amount_and_date(combined_content):
        debug_print("Skipping email due to missing dollar amount and/or date.")
        return []
    
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY must be set in .env before extracting invoice details.")

    client = OpenAI(api_key=OPENAI_API_KEY)
    prompt = (
        "You are an accountant preparing Australian tax return expense records. "
        "Use the email body and attachment text to extract invoice or receipt line items for the Australian financial year. "
        "Date format in output must conform to 'YYYY-MM-DD'. GST, Price, and Total must be rounded to 2 decimal places with a '$' sign in front. "
        "If the document shows a tax invoice number, receipt number, order number, or invoice ID, put the best identifier in 'Invoice Number'. "
        "If the supplier is clear, put it in 'Supplier'. "
        "Choose a short practical tax category such as Utilities, Software, Subscriptions, Education, Travel, Repairs/Maintenance, Office Supplies, Professional Services, or Other. "
        "Format the output as a JSON array. Each item must have 'Date', 'Category', 'Supplier', 'Invoice Number', 'Item/Product/Service', 'Price', 'GST', and 'Total'. "
        "Do not include any text other than the JSON array itself."

        "Example Output JSON Array:\n"
        "[{\n"
        "  \"Date\": \"2024-01-29\",\n"
        "  \"Category\": \"Subscriptions\",\n"
        "  \"Supplier\": \"YouTube\",\n"
        "  \"Invoice Number\": \"\",\n"
        "  \"Item/Product/Service\": \"YouTube: Watch, Listen, Stream\",\n"
        "  \"Price\": \"$10.00\",\n"
        "  \"GST\": \"$0.99\",\n"
        "  \"Total\": \"$10.99\"\n"
        "}]\n\n"

        "Now, extract the details from the following email content:\n\n"
        f"{combined_content}\n\n"
    )
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "user", "content": prompt}
        ]
    )

    # Print the JSON response
    debug_print(f"JSON response: {response}")

    details = response.choices[0].message.content.strip()
        
    # Ensure details are correctly formatted and extracted
    try:
        details_json = parse_json_array(details)
        debug_print(f"Extracted JSON details: {details_json}")
    except (json.JSONDecodeError, ValueError) as e:
        debug_print(f"JSON decode error: {e}")
        return []

    debug_print(f"Model response: {details}")
    if getattr(response, 'usage', None):
        debug_print(f"Tokens used: {response.usage.total_tokens}")

    return details_json

def first_present(row, *keys):
    for key in keys:
        value = row.get(key)
        if value not in (None, ''):
            return str(value).strip()
    return ''

def normalize_text(value):
    value = str(value or '').lower()
    value = re.sub(r'[^a-z0-9]+', ' ', value)
    return re.sub(r'\s+', ' ', value).strip()

def normalize_amount(value):
    if value in (None, ''):
        return ''
    match = re.search(r'-?\d+(?:,\d{3})*(?:\.\d+)?', str(value))
    if not match:
        return ''
    return f"{float(match.group(0).replace(',', '')):.2f}"

def amount_as_float(value):
    normalized = normalize_amount(value)
    return float(normalized) if normalized else None

def normalize_date(value):
    value = str(value or '').strip()
    if not value:
        return ''

    known_formats = (
        '%Y-%m-%d',
        '%d/%m/%Y',
        '%d-%m-%Y',
        '%d.%m.%Y',
        '%d %b %Y',
        '%d %B %Y',
        '%b %d, %Y',
        '%B %d, %Y',
    )
    for date_format in known_formats:
        try:
            return datetime.strptime(value, date_format).strftime('%Y-%m-%d')
        except ValueError:
            continue
    return value

def normalize_invoice_number(value):
    value = str(value or '').strip().upper()
    value = re.sub(
        r'^(?:TAX\s+)?(?:INVOICE|RECEIPT|ORDER)\s*(?:NUMBER|NO\.?|ID)?[\s:#-]+',
        '',
        value,
    )
    value = re.sub(r'^(?:ID|NUMBER|NO\.?|#)[\s:#-]+', '', value)
    return re.sub(r'[^A-Z0-9-]+', '', value)

def get_subject(msg):
    if isinstance(msg, dict) and msg.get('source') == 'graph':
        return msg.get('subject') or ''
    return decode_mime_words(msg['subject'] if msg['subject'] else '')

def build_source_metadata(msg, attachments):
    attachment_names = [attachment['filename'] for attachment in attachments]
    if isinstance(msg, dict) and msg.get('source') == 'graph':
        return {
            'Source Subject': msg.get('subject') or '',
            'Source Message ID': msg.get('message_id') or '',
            'Source Date': msg.get('date') or '',
            'Source Attachments': '; '.join(attachment_names),
        }

    raw_date = msg.get('Date', '')
    source_date = ''
    try:
        parsed_date = email.utils.parsedate_to_datetime(raw_date)
        source_date = parsed_date.isoformat()
    except (TypeError, ValueError):
        source_date = raw_date or ''

    return {
        'Source Subject': decode_mime_words(msg['subject'] if msg['subject'] else ''),
        'Source Message ID': msg.get('Message-ID', ''),
        'Source Date': source_date,
        'Source Attachments': '; '.join(attachment_names),
    }

def attach_source_metadata(rows, source_metadata):
    for row in rows:
        if isinstance(row, dict):
            row.update(source_metadata)
    return rows

def row_signature(row, include_item=True):
    parts = [
        normalize_date(first_present(row, 'Date', 'Datetime')),
        normalize_amount(row.get('Total', '')),
        normalize_invoice_number(row.get('Invoice Number', '')),
        normalize_text(row.get('Supplier', '')),
    ]
    if include_item:
        parts.append(normalize_text(row.get('Item/Product/Service', '')))
    return tuple(parts)

def duplicate_candidates(row):
    date = normalize_date(first_present(row, 'Date', 'Datetime'))
    total = normalize_amount(row.get('Total', ''))
    item = normalize_text(row.get('Item/Product/Service', ''))
    supplier = normalize_text(row.get('Supplier', ''))
    invoice_number = normalize_invoice_number(row.get('Invoice Number', ''))
    source_message_id = normalize_text(row.get('Source Message ID', ''))
    source_attachments = normalize_text(row.get('Source Attachments', ''))

    candidates = []
    if invoice_number and total:
        candidates.append((
            f"invoice-total:{invoice_number}:{total}",
            f"same invoice/receipt number ({invoice_number}) and total (${total})",
        ))
    if date and total and item:
        candidates.append((
            f"date-total-item:{date}:{total}:{item}",
            f"same date ({date}), total (${total}), and item/service",
        ))
    if source_message_id and total and item:
        candidates.append((
            f"message-total-item:{source_message_id}:{total}:{item}",
            "same source message, total, and item/service",
        ))
    if not candidates and total and item and supplier:
        candidates.append((
            f"supplier-total-item:{supplier}:{total}:{item}",
            f"same supplier, total (${total}), and item/service",
        ))
    if source_attachments and total and item:
        digest = hashlib.sha1(source_attachments.encode('utf-8')).hexdigest()[:12]
        candidates.append((
            f"attachments-total-item:{digest}:{total}:{item}",
            "same attachment set, total, and item/service",
        ))

    return candidates

def flag_possible_duplicates(rows):
    key_groups = {}
    for index, row in enumerate(rows):
        if not isinstance(row, dict):
            continue
        row['Date'] = normalize_date(first_present(row, 'Date', 'Datetime'))
        for key, reason in duplicate_candidates(row):
            key_groups.setdefault(key, {'reason': reason, 'indexes': []})['indexes'].append(index)

    duplicate_groups = []
    seen_group_indexes = set()
    for key, data in key_groups.items():
        indexes = data['indexes']
        if len(indexes) < 2:
            continue
        frozen_indexes = tuple(indexes)
        if frozen_indexes in seen_group_indexes:
            continue
        seen_group_indexes.add(frozen_indexes)
        duplicate_groups.append((key, data['reason'], indexes))

    for group_number, (key, reason, indexes) in enumerate(duplicate_groups, start=1):
        duplicate_of = indexes[0] + 2  # CSV row number, accounting for the header row.
        for index in indexes:
            row = rows[index]
            row['Possible Duplicate'] = 'Yes'
            row['Duplicate Group'] = f'DUP-{group_number:03d}'
            row['Duplicate Reason'] = reason
            row['Possible Duplicate Of'] = '' if index == indexes[0] else str(duplicate_of)

    for row in rows:
        if not isinstance(row, dict):
            continue
        row.setdefault('Possible Duplicate', 'No')
        row.setdefault('Duplicate Group', '')
        row.setdefault('Duplicate Reason', '')
        row.setdefault('Possible Duplicate Of', '')

    return rows

def write_to_csv(data, csv_path):
    with open(csv_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=CSV_FIELDS, quoting=csv.QUOTE_ALL, extrasaction='ignore')
        writer.writeheader()
        for row in data:
            if isinstance(row, dict):
                debug_print(f"Writing row to CSV: {row}")
                writer.writerow({field: row.get(field, '') for field in CSV_FIELDS})
            else:
                debug_print(f"Invalid row detected, not a dictionary: {row}")

def should_remove_duplicate(row, first_row):
    row_total = amount_as_float(row.get('Total', ''))
    first_total = amount_as_float(first_row.get('Total', ''))
    row_price = amount_as_float(row.get('Price', ''))
    first_price = amount_as_float(first_row.get('Price', ''))

    same_line_item = row_signature(row, include_item=True) == row_signature(first_row, include_item=True)
    same_invoice_total = row_signature(row, include_item=False) == row_signature(first_row, include_item=False)
    same_source = (
        normalize_text(row.get('Source Message ID', ''))
        and normalize_text(row.get('Source Message ID', '')) == normalize_text(first_row.get('Source Message ID', ''))
    )
    same_attachment = (
        normalize_text(row.get('Source Attachments', ''))
        and normalize_text(row.get('Source Attachments', '')) == normalize_text(first_row.get('Source Attachments', ''))
    )

    # Remove exact repeated line items. For same-invoice rows with different item text,
    # only remove when the price also equals the invoice total, which usually means the
    # whole invoice was extracted twice rather than separate line items on one receipt.
    if same_line_item:
        return True
    if same_invoice_total and same_source:
        return True
    if same_invoice_total and same_attachment:
        return True
    if same_invoice_total and row_price == row_total and first_price == first_total:
        return True
    return False

def deduplicate_rows(rows):
    deduped = []
    group_first_rows = {}
    for row in rows:
        if not isinstance(row, dict):
            continue

        duplicate_group = row.get('Duplicate Group', '')
        if row.get('Possible Duplicate') != 'Yes' or not duplicate_group:
            deduped.append(row)
            continue

        first_row = group_first_rows.setdefault(duplicate_group, row)
        if row is first_row:
            deduped.append(row)
            continue

        if should_remove_duplicate(row, first_row):
            debug_print(f"Removing duplicate row from {duplicate_group}: {row}")
            continue
        deduped.append(row)

    return deduped

def main():
    initialize_output_files()
    results = []
    emails = fetch_emails()
    filtered_emails = [
        msg for msg in emails
        if any(keyword in get_subject(msg).lower() for keyword in SUBJECT_KEYWORDS)
    ]

    debug_print(f"Total emails passing the filter: {len(filtered_emails)}")
    
    for msg in filtered_emails:
        subject = get_subject(msg)
        body = get_body(msg)
        attachments = get_attachments(msg)
        
        debug_print(f"Processing email with subject: {subject}")
        debug_print(f"Body size: {len(body) if body else 'No body'}")
        debug_print(f"Number of attachments: {len(attachments)}")
        
        details = extract_details_from_content(subject, body, attachments)
        details = attach_source_metadata(details, build_source_metadata(msg, attachments))
        debug_print(f"Extracted details: {details}")
        
        if details:
            results.extend(details)

    results = flag_possible_duplicates(results)
    write_to_csv(results, csv_path)
    deduped_results = deduplicate_rows(results)
    write_to_csv(deduped_results, dedup_csv_path)
    print(f"Wrote {len(results)} rows to {csv_path}")
    print(f"Wrote {len(deduped_results)} rows to {dedup_csv_path}")

if __name__ == "__main__":
    main()
