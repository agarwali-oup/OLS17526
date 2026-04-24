import os
import re
import shutil
import zipfile
from html import unescape
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

import pandas as pd
import pdfplumber
import requests


# ============================================================
# CONFIG
# ============================================================
INPUT_FILE = r"C:\Users\agarwais\Downloads\JiraTasks\539_577.xlsx"
OUTPUT_FILE = r"C:\Users\agarwais\Downloads\JiraTasks\539_577_filled.xlsx"
DOWNLOAD_DIR = r"C:\Users\agarwais\Downloads\JiraTasks"

# Paste your browser cookie string here exactly as copied
# (raw string is fine; code will also fix &amp; automatically)
MANUAL_COOKIE = r"""SPHomeWeb:AppBar/useLegacy=false; SIMI=eyJzdCI6MTc3MTIxOTUwOTg2Nn0=; WordWacDataCenter=GEU5; WordWacDataCenterSetTime=2026-04-16T08:05:57.546Z; ScaleCompatibilityDeviceId=8f7cbbf0-b663-4295-9718-86879789f0ae; SPA_RT=eyJzcGFfcnQiOiIxLkFYUUFZaHQya1VWTTlVT2ZEcjZLMmJWUl8zYUk0UWgzWVg1SXVMWFBsUXdlV1l3X0FiSjBBQS5CUUFCQXdFQUFBQURBT3pfQlFEMF8wVjJiMU4wYzBGeWRHbG1ZV04wY3dJQUFBQUFBRzRIWmRkNjdIZnZGTkNSMzhpOW9GTVVtRVppZF82S2ZiMnkyaHBONS1LT1BFMkkyVnItTGZJOFYxY05mTHlINzM2Rm1ON1JWOENYa2dDeGd2Y1ZiOFI5YmJqUnU3Z3U1OE82UnFGQUllRzZEdU83d0NleUp3TkUtVS1mbjBydjVMLUtNOFM2aUVuaXUtTEtCWUtMY3F1SkVKV3hJck0zU25rMDhtX1ZjR2huOGktV1pJX2VtQjlLajdFQllOTUVJeTBXNnhNckJ5UWVVR3BhWTNiSi05WDU5aXp2UnJiU0NlQlhoS1JwM0xJd1djOGxVaUgxTXdUN1dZdWppaWp1WnNUcGhLaGQxb1g1SGdfcDNTdUNhbnRfR0hxNVBqeUtVVFo4dndYaG5QcDBOYmpNbndIVllhRmRjR2lOZGtYdHhueFo4RG5hei1wZElQNWZBQi12RWlZTEs4enFmQ1RjdFE4N3AxcmkzX0kteGQyQ2hlNEZGRmd3Z3Rqbmhwbnd2NXNzXzNwcDZRTlFUNVE0Q0R0d1pjLUd6czBKd1RZVEVUX3BpZXJydXFybkhqV1o3a1ZSYnhmY09Nc09NNUVKUXc2RGtKbDhwMFhlWk90SWNCbXNXSDZQVjM2TnJ6cnJ3S3BFUGpYV3lITEFGN0tvMDFfNW1uWDVWaGstaktCLWEtNGRHd0hvaV9DdFFneXpQRzNNLXRKb3R5UW9XcE1zRkdfaURQVlRZMlRNOHpMQUZaN2hvNkhUNHZOMlJxTEZrbFdCaS1ZcGFYNkszLWRPeTZZTkRWcURwem1BckFGNFV1Z3ZlQmNFT3huSTFJQ3I0TXdCaC1uMkJGdGhEb3AzZ3VodmgwX3kzVnhiRk1feGw4ZVFkYnlNOUN2OTFXNHlqS0pLV3Q4TFl6RDBXU3UxOUI2NnNrSWo0OFNxN0RwcUF4VFk5ZnVSclBPVmREYkhMVkxTZEQ2YlNOWWVIN0wwSnJsYWlsUXdReS1rUTR2QmlkYkJBTDYxeEV2MG1QVGVOX2d6ejI3cmJxZzE4aVVBajlyb3gtV2N4OW1HSjJ3U2V1VUthdnZnM2VFUERxaDlQVnFJQUV0OWlyNzg3UkszUk0xSUNKWjlicU5ManRpOGdFR2EyN0dLSFYyS2hRd1hFaFYxTUtnRFRXZzREc3RiV3oycDRVUGxZVF9aa29zWmxBZk5Mb1E5ZVMwTDFUcEtTSVhmWVFRRndlRlF1ZGp4aFdSalhXYUpkT0VyaWU4cW1sQmFuQXRBemtUeWl3Y1lfVjcxaFRKbGp1bHlNWnYzMl9TQ1V3TGVTSmp1aWVDQUxuWDFjQXQtZWdUWmpISk5WaGdqMU5RaTBuRlkwbGdqSHZrb1M5ZTBvZlA5MWplajQxTjUzbFY5dzJFWmx4Wi15VFA1NVVBeFg4OWZNcEh5QlNObVotLWhxbklXcS1sSWQ0R1lmRjFRNk9WaUxZSkM5RTIyR0pTY082b3RhZ3hCQzJuTUJPb29kR3NJVEZHbWdZcnZDa0s0ZFNBaHFlbEtuNFkwUGxZOXp6Yk9QeUpPdmVIdXBHOVFSWDF1b3Y5cU9reXBIalNweWU1UmlKZkV3X1BvdXl1WDZRZGI3RVF0QUhLd1pIak1Wb1VQWUtOZmlMMndKLXp5VjdRVFYtZU5aWmxqajVBM1hXaUZoSFQtNG9JcGVrYzRZSWJFOUJ1a0VXRWNla1dRbFB2azBha3pwVmRkcXBzT25zRzdmODNvNlpvZUUzMENubTNhQWlRWDdsWEpDV3c1d09wUjBBVS1SMHYwbmZwaG9YenJuazR0aVdOd20zN09GQnR0aDY4OGRkUTNPVGo1alFhV0NLdFhrdUNHaDRuY0plMFJEeTVadlljR0o5R2ZJNnFMVWNkM3JQUjc4Y2ZyZEZZRkViWklLam1OVzd1Wk15RHRVZmRhd0pWTWRYQ1hPNVR5a09rNmtWQmk0NGpvWjd0dFZhMGNBOElsMnZ3STFzR3cwMXVaUlVSWkNpTGxsYTkzVDM1NDFIYjlRbGxYM0hCeHIyX1l2LUh5YWtvWHpXa0sxaHMyMmVEMXNDUXYxN05pRWxOMVJJS3ctSmhIOGMwSGNiOGpxUHFjVmxucDZwdGFwSUc2dGo5Yi0zRlBWYkFYU0hRYXpUa2JBX0lOWUpRd2IzREhaRVJSdFQ3OTN2VVg4SXF1SGMtNzZsNG15VWlHOU1uZE4yb250XzV6dHdpRDhrSjNmeGxndGlyVWo5cEF4VXhNTWRNSkdiMUtVZmtGTGdrR1dNV0IzdGFucUdFTWo1eEdUcUk1NWpja3V4UHFqZyIsIm9pZCI6ImQ3NzIxNjg4LTc2MmYtNDY4ZC05NTM4LWRlMTY1NzMzYjMxMCIsImNsaWVudF9pbmZvIjoiZXlKMWFXUWlPaUprTnpjeU1UWTRPQzAzTmpKbUxUUTJPR1F0T1RVek9DMWtaVEUyTlRjek0ySXpNVEFpTENKMWRHbGtJam9pT1RFM05qRmlOakl0TkdNME5TMDBNMlkxTFRsbU1HVXRZbVU0WVdRNVlqVTFNV1ptSW4wIiwic3BhX3JlZnJlc2hfdG9rZW5fZXhwIjoiMjAyNi0wNC0yNVQwNjoxNzowNy41NjMyMjI4WiJ9; rtFa=WKH6UvwT+ukqj2AfqjX0ZHW6vD8JEi7HbPfD2IdoFqImOTE3NjFiNjItNGM0NS00M2Y1LTlmMGUtYmU4YWQ5YjU1MWZmIzEzNDIxMTM5NDI4NjI0Njc1MSM1ZmJiMGJhMi0wMDQ0LWMwMDAtYjgxZS00ZGMwYWMzOTE4OGEjSXNoaXRhLkFnYXJ3YWwlNDBvdXAuY29tIzIwNTE3MCNTN2Vkek16bUNJWXFTUjVFR0c3RGR1bVA4RFUjUzdlZHpNem1DSVlxU1I1RUdHN0RkdW1QOERVU9C/TbTCo1VW9DrIZ5tEYyJqyjBQmi/T7hBLVmQb7vbBdpMUWh7Vx207zyhRK3Zb+Bt7b7UnL2FpjTrjemLn4Avrq+1saVxaGVE+3tc3FGjFT1ZNRmSNrTjS86h2A/21tm6wrIPfIlnhwz+rdZaamD+bPMcc4Itx6Njpx+JV0GWJ44DpRoOoLXUdmYkh9ahunEYiXBqD55vPRqxHl+YeWyA3cJ7ZAABPLp7cLxT4tIBQGi4Ar2nzfRSNxM2PhEKWG1chjLtnXVwfS/jpi35dh5i7i/ujvNwjb0kk1H0ut3u3+MNliqLbURfZwtJno07oxKFEk2QaCOzp5ldKlWJzOdUAAAA=; FedAuth=77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48U1A+VjE1LDBoLmZ8bWVtYmVyc2hpcHwxMDAzMjAwNGQ4Yjc5YzFkQGxpdmUuY29tLDAjLmZ8bWVtYmVyc2hpcHxpc2hpdGEuYWdhcndhbEBvdXAuY29tLDEzNDIxMTM3NDkzMDAwMDAwMCwxMzQyMTEzMjMyODAwMDAwMDAsMTM0MjE1ODEzNTE1MTc3MTM3LDI0MDE6NDkwMDpiZTgyOjJiYjk6OWMyMjo3MzFmOmYxYTk6YmM2OSwxOTUsOTE3NjFiNjItNGM0NS00M2Y1LTlmMGUtYmU4YWQ5YjU1MWZmLCwwMDQwMWUyYS0zMGE4LTZjZGMtZTYyNC03ZDQ4Mjk2NDcxYjYsNWZiYjBiYTItMDA0NC1jMDAwLWI4MWUtNGRjMGFjMzkxODhhLGQ1YzQwYmEyLTkwZTUtYzAwMC1iODFlLTQ4Yzc2MjA2MWJhMCwsMCwxMzQyMTE1Mjk1MTUwNTU4NTgsMTM0MjE0MDg1NTE1MDU1ODU4LCwsZXlKallYQnZiR2xrYzE5c1lYUmxZbWx1WkNJNklsdGNJbU0wWmpGaVpUUmlMVFV5TW1NdE5HRmtNaTFoTUdVMExUZ3dZVEl3WkRaak0yRTJPRndpTEZ3aU9EQXpOakZqWldZdFpqWXlaUzAwWWpsaExUaGxaalF0WTJRNU1URmhOR0l3WW1VNVhDSmRJaXdpZUcxelgyTmpJam9pVzF3aVExQXhYQ0pkSWl3aWVHMXpYM056YlNJNklqRWlMQ0p3Y21WbVpYSnlaV1JmZFhObGNtNWhiV1VpT2lKSmMyaHBkR0V1UVdkaGNuZGhiRUJ2ZFhBdVkyOXRJaXdpZFhScElqb2lNemM0TWpoT1pIZzRhMjFNV0VkM2QwVlNaMEpCUVNJc0ltRjFkR2hmZEdsdFpTSTZJakV6TkRJeE1UTTNORGt6TURBd01EQXdNQ0o5LDI2NTA0Njc3NDM5OTk5OTk5OTksMTM0MjExMzk0MjcwMDAwMDAwLGQ3NzIxNjg4LTc2MmYtNDY4ZC05NTM4LWRlMTY1NzMzYjMxMCwsLCwsLDExNTI5MjE1MDQ2MDY4NDY5NzYsLDIwNTE3MCxRZWNvQ1lUbFB5NkxkQW5qbFRsdC16M3owSDAsLDIwNTE3MCxRZWNvQ1lUbFB5NkxkQW5qbFRsdC16M3owSDAsaFVaUE9zK2hqY0U1d0F0dW9qTlVIZ3hKeFhobnM1MkErRFJwa0dEdmEvbThHVGhORkZFSGdKRGZ2eWdFcThicW9mSmwxZHdhYWVrRHNVVGtLdEtPbEs5N2QzUzdaOXZnaCs1Zy82anV5TVc1MU0vNnpieHZEeTk0MjQ5ZzFiMzQwU1RWQ2kyVmVrdmx4UmdsT2VTUzNtT1UwbHQvTlZtZWxSZXlsZ2RtN2Y2clFnaWc0MHhHcCtEbEtCdkNKZzZkY3BYa0x6VzhmTmRvNW1ydUV1NnJIMk96amNHQkRWQS81eTliQks3Q1BJeVFuajhjUENRenNqZ0pRQ2VNak5QbFkrbkpXNEt6bEJmNGc0YVdCangyUVE4Zy9jZ1d0RzduNFdiUm9mU1I0bzVhVUg0YVRGbkp6czZhc1kvKzNTcWY4dlpSck1QdzA3TVlSSzFySFpCLzVBPT08L1NQPg==; FeatureOverrides_experiments=[]; msal.cache.encryption=%7B%22id%22%3A%22019daa22-006e-731b-bf38-d733efa34686%22%2C%22key%22%3A%22Pb0fBCLwS5XwKCSgP7j7DJp1P7_HbXsicnUHpqMdOrs%22%7D; SPHomeWeb:NGSP/experienceActive=false; ExcelWacDataCenter=PSG4; WacDataCenter=PSG4; ExcelWacDataCenterSetTime=2026-04-20T09:04:48.219Z; WacDataCenterSetTime=2026-04-20T09:04:48.219Z"""

os.makedirs(DOWNLOAD_DIR, exist_ok=True)


# ============================================================
# INPUT / OUTPUT COLUMN POSITIONS (0-based)
# ============================================================
# Input:
# A = 0 -> IP Number
# D = 3 -> Agent Name
# E = 4 -> URL
COL_IP = 0
COL_AGENT = 3
COL_URL = 4

# Output:
# I = 8, J = 9, K = 10, L = 11, M = 12, N = 13, O = 14
OUT_I = 8
OUT_J = 9
OUT_K = 10
OUT_L = 11
OUT_M = 12
OUT_N = 13
OUT_O = 14
OUT_P = 15
OUT_T = 19
OUT_AJ = 35
OUT_AK = 36

# ============================================================
# YOUR EXISTING HELPERS (preserved / cleaned)
# ============================================================
def extract_zip(filepath, extract_to):
    try:
        with zipfile.ZipFile(filepath, "r") as zip_ref:
            zip_ref.extractall(extract_to)
        return extract_to
    except Exception as e:
        print(f"Zip extraction error for {filepath}: {e}")
        return None


def get_all_pdfs(folder_path):
    pdf_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(".pdf"):
                pdf_files.append(os.path.join(root, file))
    return pdf_files


def extract_text_from_pdf(filepath):
    text = ""
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
    except Exception as e:
        print(f"PDF read error for {filepath}: {e}")
        return None
    return text.strip()


def is_readable(text):
    if not text or len(text) < 100:
        return False

    weird_ratio = len(re.findall(r"[^\x00-\x7F]", text)) / max(len(text), 1)
    if weird_ratio > 0.3:
        return False

    return True

import re
from difflib import SequenceMatcher


def normalize_ocr_text(text):
    """
    Normalize OCR-broken text while preserving enough structure for regex checks.
    """
    if not text:
        return ""

    text = text.upper()

    # Normalize punctuation variants often seen in extracted PDF text
    text = text.replace("’", "'").replace("‘", "'").replace("`", "'")
    text = text.replace("–", "-").replace("—", "-")

    # Collapse repeated whitespace
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def letters_only(text):
    """
    Keep only letters and digits.
    Helps recover OCR-broken strings like:
    'P BLISHING AGREEME 'T' -> 'PBLISHINGAGREEMET'
    """
    return re.sub(r"[^A-Z0-9]", "", normalize_ocr_text(text))


def fuzzy_contains_phrase(text, phrase, threshold=0.90, window_slop=2):
    """
    Precision-first fuzzy phrase matcher.
    - Uses letters-only comparison to handle OCR splits/punctuation
    - Uses high threshold to avoid false positives
    """
    clean_text = letters_only(text)
    clean_phrase = letters_only(phrase)

    if not clean_text or not clean_phrase:
        return False

    # Exact fast path
    if clean_phrase in clean_text:
        return True

    phrase_len = len(clean_phrase)
    min_len = max(4, phrase_len - window_slop)
    max_len = phrase_len + window_slop

    for win_len in range(min_len, max_len + 1):
        for i in range(0, max(0, len(clean_text) - win_len + 1)):
            chunk = clean_text[i:i + win_len]
            ratio = SequenceMatcher(None, chunk, clean_phrase).ratio()
            if ratio >= threshold:
                return True

    return False


def regex_contains(text, pattern):
    """
    Regex search on normalized OCR text.
    """
    return re.search(pattern, normalize_ocr_text(text), re.IGNORECASE) is not None

def phrase_present(text, exact_patterns=None, fuzzy_phrases=None, fuzzy_threshold=0.90):
    """
    Returns True if ANY exact pattern matches OR ANY fuzzy phrase matches.
    Uses normalized text + fuzzy fallback.
    """
    exact_patterns = exact_patterns or []
    fuzzy_phrases = fuzzy_phrases or []

    norm = normalize_ocr_text(text)

    for pattern in exact_patterns:
        if regex_contains(norm, pattern):
            return True

    for phrase in fuzzy_phrases:
        if fuzzy_contains_phrase(norm, phrase, threshold=fuzzy_threshold):
            return True

    return False

def check_contract_structure(text):
    """
    High-precision OCR-tolerant contract detection.

    Current logic preserved.
    Added new mandatory terms:
    - ROYALTIES
    - SCHEDULE III STANDARD TERMS AND CONDITIONS
    - PUBLISHER'S OBLIGATIONS
    """

    if not text:
        return "No"

    norm = normalize_ocr_text(text)

    # Since you removed segregation, use whole text
    top_text = text
    bottom_text = text

    # --------------------------------------------------
    # STEP 1: Strong anchor (REQUIRED)
    # --------------------------------------------------

    publishing_title_hit = fuzzy_contains_phrase(
        top_text,
        "PUBLISHING AGREEMENT",
        threshold=0.88,
        window_slop=2
    )

    license_title_hit = (
        fuzzy_contains_phrase(top_text, "LICENSE AGREEMENT", threshold=0.90, window_slop=2) or
        fuzzy_contains_phrase(top_text, "LICENCE AGREEMENT", threshold=0.90, window_slop=2)
    )

    agreement_dated_hit = (
        regex_contains(top_text, r"AGREEMENT\s*,?\s*DATED") or
        regex_contains(top_text, r"AGREEMENT\s+DATED\s+AS\s+OF") or
        fuzzy_contains_phrase(top_text, "AGREEMENT DATED", threshold=0.91, window_slop=2)
    )

    strong_anchor = publishing_title_hit or license_title_hit or agreement_dated_hit

    if not strong_anchor:
        return "No"

    # --------------------------------------------------
    # STEP 2: Supporting legal structure
    # --------------------------------------------------

    whereas_hit = regex_contains(norm, r"\bWHEREAS\b")

    schedule_1_hit = regex_contains(norm, r"\bSCHEDULE\s*I\b")
    schedule_2_hit = regex_contains(norm, r"\bSCHEDULE\s*II\b")
    schedule_3_hit = regex_contains(norm, r"\bSCHEDULE\s*III\b")

    schedule_count = sum([schedule_1_hit, schedule_2_hit, schedule_3_hit])

    signed_hit = (
        regex_contains(bottom_text, r"SIGNED\s+BY") or
        fuzzy_contains_phrase(bottom_text, "SIGNED BY", threshold=0.92, window_slop=1)
    )

    # --------------------------------------------------
    # STEP 2B: NEW MANDATORY TERMS
    # --------------------------------------------------

    royalties_hit = phrase_present(
        text,
        exact_patterns=[
            r"\bROYALTIES\b"
        ],
        fuzzy_phrases=[
            "ROYALTIES"
        ],
        fuzzy_threshold=0.92
    )

    if not royalties_hit:
        return "No"

    schedule_iii_terms_hit = phrase_present(
        text,
        exact_patterns=[
            r"\bSCHEDULE\s*(?:III|3)\s+STANDARD\s+TERMS\s+AND\s+CONDITIONS\b"
        ],
        fuzzy_phrases=[
            "SCHEDULE III STANDARD TERMS AND CONDITIONS"
        ],
        fuzzy_threshold=0.90
    )

    if not schedule_iii_terms_hit:
        return "No"

    publisher_obligations_hit = phrase_present(
        text,
        exact_patterns=[
            r"\bPUBLISHER'?S\s+OBLIGATIONS\b"
        ],
        fuzzy_phrases=[
            "PUBLISHER'S OBLIGATIONS"
        ],
        fuzzy_threshold=0.90
    )

    if not publisher_obligations_hit:
        return "No"

    # --------------------------------------------------
    # STEP 3: Final decision rule
    # --------------------------------------------------
    if (whereas_hit and signed_hit) or \
       (whereas_hit and schedule_count >= 3) or \
       (signed_hit and schedule_count >= 3):
        return "Yes"

    return "No"


# def check_contract_structure(text):
#     required_patterns = [
#         r".+ Agreement",
#         r"Agreement dated",
#         r"WHEREAS",
#         r"SCHEDULE I",
#         r"SCHEDULE II",
#         r"SCHEDULE III",
#         r"Signed by"
#     ]

#     for pattern in required_patterns:
#         if not re.search(pattern, text, re.IGNORECASE):
#             return "No"

#     return "Yes"

def ordered_fuzzy_proximity_match(
    text,
    terms,
    max_window=300,
    fuzzy_threshold=0.88
):
    """
    Checks whether terms appear in strict order and within a character window.
    Uses fuzzy matching on each term.
    """
    norm = normalize_ocr_text(text)
    cursor = 0

    for term in terms:
        found = False

        search_slice = norm[cursor: cursor + max_window]

        if fuzzy_contains_phrase(search_slice, term, threshold=fuzzy_threshold):
            idx = search_slice.find(letters_only(term))
            if idx == -1:
                idx = 0
            cursor += idx
            found = True

        if not found:
            return False

    return True

def detect_copyright_status(text, region_value):
    """
    Column T logic.

    region_value = value from Column M (US / UK)
    """

    if not text:
        return ""

    norm = normalize_ocr_text(text)

    # ==================================================
    # CASE B — Author / Editor owns copyright (PRIORITY)
    # ==================================================
    case_b_author = ordered_fuzzy_proximity_match(
        norm,
        ["THE COPYRIGHT", "VESTED", "AUTHOR"],
        max_window=350,
        fuzzy_threshold=0.88
    )

    case_b_editor = ordered_fuzzy_proximity_match(
        norm,
        ["THE COPYRIGHT", "VESTED", "EDITOR"],
        max_window=350,
        fuzzy_threshold=0.88
    )

    if case_b_author or case_b_editor:
        return f"{region_value} Agreement Author/Editor owns Copyright - No revision to Boiler Plate"

    # ==================================================
    # CASE A — Publisher owns copyright
    # ==================================================

    assigns_to_publisher = fuzzy_contains_phrase(
        norm,
        "ASSIGNS TO THE PUBLISHER FOR THE LEGAL TERM OF COPYRIGHT",
        threshold=0.86,
        window_slop=6
    )

    belongs_to_publisher = fuzzy_contains_phrase(
        norm,
        "THE COPYRIGHT IN THE WORK WILL BELONG TO THE PUBLISHER",
        threshold=0.88,
        window_slop=4
    )

    if assigns_to_publisher or belongs_to_publisher:
        return f"{region_value} Agreement OUP Owns Copyright"

    return ""

def extract_date(text):
    match = re.search(r"Agreement dated\s*[:\-]?\s*(.*)", text, re.IGNORECASE)
    if match:
        return match.group(1).split("\n")[0].strip()
    return ""

def get_text_lines(text):
    """
    Splits extracted PDF text into clean logical lines.
    Preserves order for 'next two lines' logic.
    """
    if not text:
        return []

    # Normalize line endings
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    # Split and clean
    lines = []
    for line in text.split("\n"):
        cleaned = re.sub(r"\s+", " ", line).strip()
        if cleaned:
            lines.append(cleaned)

    return lines

def extract_composite_works_section(text):
    """
    Column AK logic:
    - Find line containing 'Composite Works' (OCR tolerant)
    - Return the next two lines after it
    """
    if not text:
        return ""

    lines = get_text_lines(text)

    for idx, line in enumerate(lines):
        # Fuzzy match the trigger line
        if fuzzy_contains_phrase(
            line,
            "COMPOSITE WORKS",
            threshold=0.88,
            window_slop=2
        ):
            # Collect next two lines safely
            following_lines = []

            if idx + 1 < len(lines):
                following_lines.append(lines[idx + 1])
            if idx + 2 < len(lines):
                following_lines.append(lines[idx + 2])

            return " ".join(following_lines).strip()

    return ""

def extract_footer(filepath):
    try:
        with pdfplumber.open(filepath) as pdf:
            first_page = pdf.pages[0]
            words = first_page.extract_words()

            height = first_page.height
            footer_words = [w["text"] for w in words if w["top"] > height * 0.85]

            return " ".join(footer_words)

    except Exception as e:
        print(f"Footer extraction error for {filepath}: {e}")

    return ""


def check_agent(text, agent_name):
    if not agent_name or str(agent_name).strip().lower() == "nan":
        return "No"

    cleaned_text = text.replace("c/o", "").lower()
    return "Yes" if str(agent_name).strip().lower() in cleaned_text else "No"


def classify_us_uk(ip_value):
    ip_value = str(ip_value).strip()
    if ip_value.startswith("6"):
        return "US"
    elif ip_value.startswith("2"):
        return "UK"
    return ""


def phrase_present(text, exact_patterns=None, fuzzy_phrases=None, fuzzy_threshold=0.90):
    """
    Returns True if ANY exact pattern matches OR ANY fuzzy phrase matches.
    """
    exact_patterns = exact_patterns or []
    fuzzy_phrases = fuzzy_phrases or []

    norm = normalize_ocr_text(text)

    for pattern in exact_patterns:
        if regex_contains(norm, pattern):
            return True

    for phrase in fuzzy_phrases:
        if fuzzy_contains_phrase(norm, phrase, threshold=fuzzy_threshold):
            return True

    return False


def count_phrase_occurrences(text, exact_patterns=None, fuzzy_phrase=None, fuzzy_threshold=0.90, window_slop=2):
    """
    Count occurrences of a phrase.
    Strategy:
    1. Count exact regex matches first
    2. If exact count is zero, use non-overlapping fuzzy matching
    """
    exact_patterns = exact_patterns or []
    norm = normalize_ocr_text(text)

    # --------------------------------------------------
    # 1) Exact regex count first
    # --------------------------------------------------
    exact_count = 0
    for pattern in exact_patterns:
        exact_count += len(re.findall(pattern, norm, flags=re.IGNORECASE))

    if exact_count > 0:
        return exact_count

    # --------------------------------------------------
    # 2) Fuzzy count fallback
    # --------------------------------------------------
    if not fuzzy_phrase:
        return 0

    clean_text = letters_only(norm)
    clean_phrase = letters_only(fuzzy_phrase)

    if not clean_text or not clean_phrase:
        return 0

    phrase_len = len(clean_phrase)
    min_len = max(4, phrase_len - window_slop)
    max_len = phrase_len + window_slop

    count = 0
    i = 0

    while i <= len(clean_text) - min_len:
        matched = False

        for win_len in range(min_len, max_len + 1):
            if i + win_len > len(clean_text):
                continue

            chunk = clean_text[i:i + win_len]
            ratio = SequenceMatcher(None, chunk, clean_phrase).ratio()

            if ratio >= fuzzy_threshold:
                count += 1
                i += win_len  # skip ahead to avoid overlapping duplicate counts
                matched = True
                break

        if not matched:
            i += 1

    return count


def detect_editor_status(text):
    """
    Column P logic:
    - if neither phrase is present -> "No"
    - if one of the phrases is present -> "Editor"
    - if "Editor's Agreement" appears more than once -> "Multiple Editors"
    """
    if not text:
        return "No"

    # Count occurrences of "Editor's Agreement"
    editors_agreement_count = count_phrase_occurrences(
        text,
        exact_patterns=[
            r"\bEDITOR'?S\s+AGREEMENT\b"
        ],
        fuzzy_phrase="EDITOR'S AGREEMENT",
        fuzzy_threshold=0.90,
        window_slop=2
    )

    # Check presence of "The Editor will"
    editor_will_present = phrase_present(
        text,
        exact_patterns=[
            r"\bTHE\s+EDITOR\s+WILL\b"
        ],
        fuzzy_phrases=[
            "THE EDITOR WILL"
        ],
        fuzzy_threshold=0.90
    )

    if editors_agreement_count > 1:
        return "Multiple Editors"

    if editors_agreement_count >= 1 or editor_will_present:
        return "Editor"

    return "No"

# ============================================================
# DATAFRAME HELPERS
# ============================================================
def ensure_min_columns(df, min_cols):
    """Make sure df has at least min_cols columns."""
    while df.shape[1] < min_cols:
        df[f"__extra_{df.shape[1]}"] = ""
    return df


def set_cell_by_pos(df, row_pos, col_pos, value):
    df.iat[row_pos, col_pos] = value


def get_cell_by_pos(df, row_pos, col_pos):
    try:
        return df.iat[row_pos, col_pos]
    except Exception:
        return ""


def clean_folder(folder_path):
    if os.path.exists(folder_path):
        shutil.rmtree(folder_path, ignore_errors=True)
    os.makedirs(folder_path, exist_ok=True)


# ============================================================
# DOWNLOAD HELPERS
# ============================================================
def build_session():
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0"
        )
    })

    cookie_value = unescape(MANUAL_COOKIE).strip()
    if cookie_value:
        session.headers.update({"Cookie": cookie_value})
        print("Using manually pasted browser cookies.")
    else:
        print("WARNING: MANUAL_COOKIE is empty. SharePoint download will likely fail.")

    return session


def detect_file_type(content_bytes):
    """
    Returns: 'pdf', 'zip', or None
    """
    if content_bytes.startswith(b"%PDF"):
        return "pdf"
    if content_bytes.startswith(b"PK\x03\x04"):
        return "zip"
    return None


def save_binary(content, output_path):
    with open(output_path, "wb") as f:
        f.write(content)


def add_or_replace_query_param(url, key, value):
    """
    Cleanly add/replace query params (handles URLs better than string concat).
    """
    url = unescape(url.strip())
    parts = urlparse(url)
    query = parse_qs(parts.query, keep_blank_values=True)
    query[key] = [value]
    new_query = urlencode(query, doseq=True)
    return urlunparse((
        parts.scheme,
        parts.netloc,
        parts.path,
        parts.params,
        new_query,
        parts.fragment
    ))


def build_candidate_urls(original_url):
    """
    Try a few common SharePoint download variants.
    """
    original_url = unescape(original_url.strip())
    candidates = []

    # 1) Original
    candidates.append(original_url)

    # 2) Force download=1
    candidates.append(add_or_replace_query_param(original_url, "download", "1"))

    # 3) If web=1 exists, switch it to download=1
    lowered = original_url.lower()
    if "web=1" in lowered:
        candidates.append(re.sub(r"web=1", "download=1", original_url, flags=re.IGNORECASE))

    # remove duplicates while preserving order
    seen = set()
    unique_candidates = []
    for u in candidates:
        if u not in seen:
            unique_candidates.append(u)
            seen.add(u)

    return unique_candidates


def download_file(session, url, filename_base):
    """
    Attempts to download the real asset behind a SharePoint link using:
    - manually pasted cookies
    - original URL
    - download=1 variants

    Returns a local file path ending in .pdf or .zip, else None.
    """
    try:
        candidate_urls = build_candidate_urls(url)

        for attempt_no, candidate_url in enumerate(candidate_urls, start=1):
            print(f"  Download attempt {attempt_no}: {candidate_url}")

            try:
                response = session.get(candidate_url, timeout=90, allow_redirects=True)
            except Exception as e:
                print(f"    Request failed: {e}")
                continue

            print(f"    Status: {response.status_code}")
            print(f"    Final URL: {response.url}")
            print(f"    Content-Type: {response.headers.get('Content-Type', '')}")

            if response.status_code != 200:
                continue

            if not response.content or len(response.content) < 100:
                print("    Response too small to be a real file.")
                continue

            file_type = detect_file_type(response.content[:10])

            if file_type == "pdf":
                file_path = os.path.join(DOWNLOAD_DIR, f"{filename_base}.pdf")
                save_binary(response.content, file_path)
                print(f"    Saved PDF: {file_path}")
                return file_path

            if file_type == "zip":
                file_path = os.path.join(DOWNLOAD_DIR, f"{filename_base}.zip")
                save_binary(response.content, file_path)
                print(f"    Saved ZIP: {file_path}")
                return file_path

            # If HTML page came back, log it for debugging
            snippet = response.content[:500].decode("utf-8", errors="ignore").lower()
            if "<html" in snippet or "<!doctype html" in snippet:
                debug_html_path = os.path.join(DOWNLOAD_DIR, f"{filename_base}_debug_attempt_{attempt_no}.html")
                save_binary(response.content, debug_html_path)
                print(f"    Got HTML instead of file. Saved debug page: {debug_html_path}")
                continue

            print("    Content was not recognized as PDF or ZIP.")

        return None

    except Exception as e:
        print(f"Download error for {url}: {e}")
        return None


# ============================================================
# MAIN
# ============================================================
def main():
    print("Reading Excel...")
    df = pd.read_excel(INPUT_FILE, engine="openpyxl")
    
    # Convert output columns I:O to object dtype so they can store text values
    for col_pos in [OUT_I, OUT_J, OUT_K, OUT_L, OUT_M, OUT_N, OUT_O, OUT_P, OUT_T, OUT_AJ, OUT_AK]:
        col_name = df.columns[col_pos]
        df[col_name] = df[col_name].astype("object")

    # Make sure columns up to AK exist
    df = ensure_min_columns(df, OUT_AK + 1)

    session = build_session()

    for row_pos in range(len(df)):
        print(f"\nProcessing row {row_pos}...")

        url = str(get_cell_by_pos(df, row_pos, COL_URL)).strip()
        if url.lower().endswith('.pdf'):
            xyz= url 
            url = f"https://oxforduniversitypress.sharepoint.com/sites/RD4AI/Shared Documents/Sourcing and Product/Strategic Sourcing OUP Content (Contracts, Rights etc)/Lifecycle/MARCH 2026 CONTRACT CHECK PROGRAMME/OMO/../../../../../../../../:b:/r/sites/RD4AI/Shared%20Documents/Sourcing%20and%20Product/Strategic%20Sourcing%20OUP%20Content%20(Contracts,%20Rights%20etc)/Lifecycle/MARCH%202026%20CONTRACT%20CHECK%20PROGRAMME/OMO/OMO%20Contracts/{xyz}?csf=1&web=1&e=cNjqo4"
        
        ip_number = get_cell_by_pos(df, row_pos, COL_IP)
        agent_name = str(get_cell_by_pos(df, row_pos, COL_AGENT)).strip()

        # Default / fixed outputs
        set_cell_by_pos(df, row_pos, OUT_I, "Machine")
        set_cell_by_pos(df, row_pos, OUT_N, "Yes")
        set_cell_by_pos(df, row_pos, OUT_M, classify_us_uk(ip_number))

        # Default failure values unless proven otherwise
        set_cell_by_pos(df, row_pos, OUT_J, "No")
        set_cell_by_pos(df, row_pos, OUT_K, "")
        set_cell_by_pos(df, row_pos, OUT_L, "")
        # set_cell_by_pos(df, row_pos, OUT_O, "No")
        set_cell_by_pos(df, row_pos, OUT_P, "No")
        

        if not url or url.lower() == "nan":
            print("  No URL found in column E.")
            continue

        # ------------------------------------------
        # STEP 1: Download asset from SharePoint
        # ------------------------------------------
        file_path = download_file(session, url, f"row_{row_pos}")

        if not file_path:
            print("  Could not download a valid ZIP or PDF.")
            continue

        # ------------------------------------------
        # STEP 2: Build list of PDFs to process
        # ------------------------------------------
        pdf_list = []

        if file_path.lower().endswith(".pdf"):
            pdf_list = [file_path]

        elif file_path.lower().endswith(".zip"):
            extracted_folder = os.path.join(DOWNLOAD_DIR, f"extracted_{row_pos}")
            clean_folder(extracted_folder)

            extracted_path = extract_zip(file_path, extracted_folder)
            if not extracted_path:
                print("  ZIP extraction failed.")
                continue

            pdf_list = get_all_pdfs(extracted_path)

        else:
            print("  Downloaded file is neither PDF nor ZIP.")
            continue

        if not pdf_list:
            print("  No PDFs found to process.")
            continue

        # ------------------------------------------
        # STEP 3: Process PDFs with your logic
        # ------------------------------------------
        full_contract_found = False

        for pdf_file in pdf_list:
            print(f"  Checking PDF: {pdf_file}")

            text = extract_text_from_pdf(pdf_file)
            set_cell_by_pos(df, row_pos, OUT_P, detect_editor_status(text))
            set_cell_by_pos(df, row_pos, OUT_T, detect_copyright_status(text, df.iat[row_pos, OUT_M]))
            set_cell_by_pos(df, row_pos, OUT_AK, extract_composite_works_section(text))
            set_cell_by_pos(df, row_pos, OUT_AJ, "Yes" if str(get_cell_by_pos(df, row_pos, OUT_AK)).strip() else "No")
            #input(text[:500])  # Debug: show first 500 chars of extracted text
            if not is_readable(text):
                print("    PDF not readable enough.")
                continue

            if check_contract_structure(text) == "Yes":
                full_contract_found = True
                print("    Full contract found.")

                set_cell_by_pos(df, row_pos, OUT_J, "Yes")
                set_cell_by_pos(df, row_pos, OUT_K, extract_footer(pdf_file))
                set_cell_by_pos(df, row_pos, OUT_L, extract_date(text))
                # set_cell_by_pos(df, row_pos, OUT_O, check_agent(text, agent_name))
                # set_cell_by_pos(df, row_pos, OUT_P, detect_editor_status(text))
                # set_cell_by_pos(df, row_pos, OUT_T, detect_copyright_status(text, df.iat[row_pos, OUT_M]))
                
                break

        if not full_contract_found:
            print("  No full contract found in available PDFs.")

    # ------------------------------------------
    # SAVE OUTPUT
    # ------------------------------------------
    print("\nSaving output...")
    df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    print(f"✅ Done. Output saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()