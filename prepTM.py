import os
import re
import io
from datetime import date
from email import policy
from email.parser import BytesParser
from typing import List, Dict, Any
import csv

try:
    import extract_msg
except ImportError:
    extract_msg = None

__all__ = [
    "extract_som_entries",
]

# Hard‑coded template TYPES for Trend Micro SOM
TEMPLATE_TYPES: List[str] = [
    "domain",
    "url",
    "ip",  # IPv4
    "ip",  # IPv6
    "sha1",
    "sha256",
    "email_sender",
]


def parse_message(path: str) -> Any:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".eml":
        with open(path, 'rb') as f:
            return BytesParser(policy=policy.default).parse(f)
    elif ext == ".msg":
        if extract_msg is None:
            raise RuntimeError("extract_msg library not installed; cannot parse .msg files")
        msg = extract_msg.Message(path)
        msg.convert()
        return msg
    else:
        raise ValueError(f"Unsupported file type: {ext}")


def extract_csv_from_message(msg) -> List[Dict[str, str]]:
    # returns list of dict rows
    # .eml
    if hasattr(msg, 'walk'):
        for part in msg.walk():
            fn = part.get_filename() or ""
            if part.get_content_disposition() == 'attachment' and fn.lower().endswith('.csv'):
                raw = part.get_payload(decode=True).decode('utf-8', errors='replace')
                reader = csv.DictReader(io.StringIO(raw))
                return [row for row in reader]
    # .msg
    if extract_msg and isinstance(msg, extract_msg.Message):
        for att in msg.attachments:
            fn = att.longFilename or att.shortFilename or ""
            if fn.lower().endswith('.csv'):
                data = att.data.decode('utf-8', errors='replace')
                reader = csv.DictReader(io.StringIO(data))
                return [row for row in reader]
    return []


def extract_urls_from_message(msg) -> List[str]:
    if hasattr(msg, 'get_body'):
        html_part = msg.get_body(preferencelist=('html',))
        html = html_part.get_content() if html_part else ""
    elif extract_msg and isinstance(msg, extract_msg.Message):
        html = msg.htmlBody or ""
    else:
        html = ""
    urls = re.findall(r'h(?:xx)?tp[s]?://[^\s"<>]+', html)
    return [u.replace('hxxp', 'http') for u in urls]


def clean_values(vals: List[str]) -> List[str]:
    return [v.replace('[', '').replace(']', '').strip() for v in vals]


def format_som_entries(rows: List[Dict[str, str]], urls: List[str], desc_prefix: str = None) -> List[Dict[str, str]]:
    if desc_prefix is None:
        desc_prefix = f"cert {date.today().isoformat()}"
    # collect by type
    extracted: Dict[str, List[str]] = {
        "domain": clean_values([r["domain"] for r in rows if r.get("domain")]),
        "ip": clean_values([r["IP"] for r in rows if r.get("IP")]),
        "sha1": clean_values([r["sha1"] for r in rows if r.get("sha1")]),
        "sha256": clean_values([r["sha256"] for r in rows if r.get("sha256")]),
        "email_sender": clean_values([r["email_sender"] for r in rows if r.get("email_sender")]),
        "url": clean_values(urls),
    }
    # build output list of dicts
    output = []
    for t in TEMPLATE_TYPES:
        for val in extracted.get(t, []):
            output.append({
                "Type": t,
                "Object": val,
                "Description": desc_prefix
            })
    return output


def extract_som_entries(path: str, desc_prefix: str = None) -> List[Dict[str, str]]:
    msg = parse_message(path)
    csv_rows = extract_csv_from_message(msg)
    urls = extract_urls_from_message(msg)
    return format_som_entries(csv_rows, urls, desc_prefix)
