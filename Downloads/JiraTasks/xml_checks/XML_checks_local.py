import os
from lxml import etree
from langdetect import detect
from io import BytesIO

# ---------- CONFIG ----------
LOCAL_FOLDER_PATH = "C:/Users/agarwais/Downloads/karger_journal_article"   # path to your local XML folder
EXPECTED_LANGUAGE = "en"  # en, fr, de etc.

# Path to local JATS DTD file
DTD_PATH = "C:/Users/agarwais/Downloads/JATS-journalpublishing1.dtd"

# Map user-friendly language names -> ISO codes
LANG_MAP = {
    "english": "en",
    "french": "fr",
    "german": "de"
}


# ---------- FUNCTIONS ----------

def fetch_xml_files_from_local(folder):
    files = []
    for root, _, filenames in os.walk(folder):
        for filename in filenames:
            if filename.endswith(".xml"):
                full_path = os.path.join(root, filename)
                files.append(full_path)
    return files


def get_file_content(file_path):
    with open(file_path, "rb") as f:
        return f.read()


def is_content_valid(content):
    return content is not None and len(content.strip()) > 0


def detect_language(content):
    try:
        text = content.decode("utf-8", errors="ignore")
        return detect(text)
    except:
        return "unknown"


def validate_jats(xml_content):
    try:
        parser = etree.XMLParser(load_dtd=True, no_network=False)
        tree = etree.parse(BytesIO(xml_content), parser)

        with open(DTD_PATH, "rb") as f:
            dtd = etree.DTD(f)

        is_valid = dtd.validate(tree)
        return is_valid, dtd.error_log.filter_from_errors()

    except Exception as e:
        return False, str(e)


def process_files(folder, language_name):
    expected_lang_code = LANG_MAP.get(language_name.lower())

    if not expected_lang_code:
        print(f"Unsupported language: {language_name}")
        return

    files = fetch_xml_files_from_local(folder)

    print(f"Found {len(files)} XML files\n")

    for file_path in files:
        print(f"Processing: {file_path}")

        content = get_file_content(file_path)

        # ---- CHECK 1: Content exists ----
        if not is_content_valid(content):
            print("❌ File is empty or has no content\n")
            continue

        # ---- CHECK 2: Language ----
        detected_lang = detect_language(content)
        if detected_lang != expected_lang_code:
            print(f"❌ Language mismatch: Expected={expected_lang_code}, Found={detected_lang}")
        else:
            print("✅ Language OK")

        # ---- CHECK 3: JATS DTD Validation ----
        is_valid, errors = validate_jats(content)

        if is_valid:
            print("✅ JATS DTD validation passed")
        else:
            print("❌ JATS validation failed")
            print(errors)

        print("-" * 50)


# ---------- RUN ----------
if __name__ == "__main__":
    process_files(
        folder=LOCAL_FOLDER_PATH,
        language_name="english"
    )