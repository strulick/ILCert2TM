import re
import io
from datetime import date
from email import policy
from email.parser import BytesParser
from typing import List, Dict, Any
import pandas as pd

__all__ = [
    "extract_som_entries",
    "parse_eml",
    "extract_csv_from_eml",
    "extract_urls_from_eml",
    "clean_values",
    "format_som_entries",
]

# Hard‑coded template TYPES for Trend Micro Suspicious Object Management
TEMPLATE_TYPES = [
    "domain",
    "url",
    "ip",        # IPv4 / IPv6 entries handled by order
    "ip",
    "sha1",
    "sha256",
    "email_sender",
]


def parse_eml(eml_path: str) -> Any:
    """
    Parse an .eml file and return the email.message.Message object.
    """
    with open(eml_path, 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)
    return msg


def extract_csv_from_eml(msg: Any) -> pd.DataFrame:
    """
    Find the first .csv attachment in the email and return it as a DataFrame.
    """
    for part in msg.walk():
        filename = part.get_filename() or ""
        if part.get_content_disposition() == 'attachment' and filename.lower().endswith('.csv'):
            raw = part.get_payload(decode=True).decode('utf-8', errors='replace')
            return pd.read_csv(io.StringIO(raw))
    return pd.DataFrame()


def extract_urls_from_eml(msg: Any) -> List[str]:
    """
    Extract all URLs from the HTML body of an email, fixing `hxxp` -> `http`.
    """
    html_part = msg.get_body(preferencelist=('html',))
    html = html_part.get_content() if html_part else ""
    urls = re.findall(r'h(?:xx)?tp[s]?://[^\s"<>]+', html)
    return [re.sub(r'hxxp', 'http', u) for u in urls]


def clean_values(values: List[Any]) -> List[str]:
    """
    Convert each value to string, remove brackets, and strip whitespace.
    """
    cleaned = []
    for v in values:
        s = str(v).replace('[', '').replace(']', '').strip()
        cleaned.append(s)
    return cleaned


def format_som_entries(
    csv_df: pd.DataFrame,
    urls: List[str],
    description_prefix: str = None
) -> pd.DataFrame:
    """
    Build a DataFrame of Trend Micro SOM entries with columns:
      - Type
      - Object
      - Description (cert+YYYY-MM-DD by default)

    One row per value in TEMPLATE_TYPES; multiple values produce multiple rows.
    """
    # Determine description
    if description_prefix is None:
        today = date.today().isoformat()
        description_prefix = f"cert+{today}"

    # Gather raw extracted values
    extracted: Dict[str, List[str]] = {
        'domain':       clean_values(csv_df['domain'].dropna().tolist()) if 'domain' in csv_df.columns else [],
        'url':          clean_values(urls),
        'ip':           clean_values(csv_df['IP'].dropna().tolist()) if 'IP' in csv_df.columns else [],
        'sha1':         clean_values(csv_df['sha1'].dropna().tolist()) if 'sha1' in csv_df.columns else [],
        'sha256':       clean_values(csv_df['sha256'].dropna().tolist()) if 'sha256' in csv_df.columns else [],
        'email_sender': clean_values(csv_df['email_sender'].dropna().tolist()) if 'email_sender' in csv_df.columns else [],
    }

    # Build SOM entries, maintaining order in TEMPLATE_TYPES
    rows = []
    for t in TEMPLATE_TYPES:
        for value in extracted.get(t, []):
            rows.append({
                'Type': t,
                'Object': value,
                'Description': description_prefix
            })

    return pd.DataFrame(rows)


def extract_som_entries(eml_path: str, description_prefix: str = None) -> pd.DataFrame:
    """
    High-level function: given path to .eml, parse and return formatted
    DataFrame ready for Trend Micro Suspicious Object Management import.
    """
    msg = parse_eml(eml_path)
    csv_df = extract_csv_from_eml(msg)
    urls = extract_urls_from_eml(msg)
    return format_som_entries(csv_df, urls, description_prefix)