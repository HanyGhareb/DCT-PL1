# ==========================================
# DCT DOCUMENT CLASSIFIER — GOOGLE COLAB
# ==========================================
# Classifies documents received from events, parties, or vendor
# engagements into: Contracts | P&L | Unrelated
#
# INPUT : A folder path (supports recursive sub-folders)
# OUTPUT: Output_{YYYYMMDD_HHMMSS}/ with 3 organised sub-folders
#
# SUPPORTED FILE TYPES: PDF, DOCX, TXT
# ==========================================


# ── STEP 1: Install dependencies ──────────────────────────────
print("Installing dependencies (this only runs once)...")
import subprocess
subprocess.run(["pip", "install", "-q", "PyMuPDF", "python-docx", "requests"], check=False)

import os, re, json, time, shutil, requests
from datetime import datetime
from pathlib import Path

try:
    import fitz          # PyMuPDF — PDF text extraction
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False
    print("WARNING: PyMuPDF not installed — PDF files will be classified by filename only")

try:
    from docx import Document as DocxDocument
    DOCX_SUPPORT = True
except ImportError:
    DOCX_SUPPORT = False
    print("WARNING: python-docx not installed — DOCX files will be classified by filename only")


# ── STEP 2: Configuration  ← EDIT THESE ──────────────────────
# Path to the folder containing your documents (with optional sub-folders)
INPUT_FOLDER  = "/content/drive/MyDrive/DCT_Documents"

# Where to create the Output_{timestamp} folder
OUTPUT_PARENT = "/content/drive/MyDrive"

# Gemini API key — set this in Colab Secrets (key name: GOOGLE_API_KEY)
# or paste directly: GEMINI_API_KEY = "AIzaSy..."
try:
    from google.colab import userdata
    GEMINI_API_KEY = userdata.get('GOOGLE_API_KEY')
except Exception:
    GEMINI_API_KEY = ""   # ← paste key here if not using Colab Secrets

GEMINI_MODEL        = "gemini-1.5-flash-latest"   # fast & free tier
SUPPORTED_EXTS      = {'.pdf', '.docx', '.doc', '.txt'}
MAX_TEXT_CHARS      = 5000   # chars sent to AI (enough for classification)
DELAY_BETWEEN_FILES = 1.5    # seconds between API calls (respect free-tier limit)


# ── STEP 3: Mount Google Drive ────────────────────────────────
try:
    from google.colab import drive
    drive.mount('/content/drive')
    print("Google Drive mounted.\n")
except Exception:
    print("Google Drive not mounted (running outside Colab or already mounted).\n")


# ── STEP 4: Text extraction helpers ───────────────────────────

def extract_text_from_pdf(path: Path) -> str:
    if not PDF_SUPPORT:
        return ""
    try:
        doc  = fitz.open(str(path))
        text = "".join(page.get_text() for page in doc)
        doc.close()
        return text
    except Exception as e:
        print(f"  ⚠  PDF read error: {e}")
        return ""


def extract_text_from_docx(path: Path) -> str:
    if not DOCX_SUPPORT:
        return ""
    try:
        doc = DocxDocument(str(path))
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception as e:
        print(f"  ⚠  DOCX read error: {e}")
        return ""


def extract_text_from_txt(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="ignore")
    except Exception as e:
        print(f"  ⚠  TXT read error: {e}")
        return ""


def extract_text(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".pdf":
        return extract_text_from_pdf(path)
    elif ext in {".docx", ".doc"}:
        return extract_text_from_docx(path)
    elif ext == ".txt":
        return extract_text_from_txt(path)
    return ""


# ── STEP 5: Gemini classification ─────────────────────────────

CATEGORIES = ["Contracts", "P&L", "Unrelated"]

def classify_document(filename: str, text: str, retries: int = 2):
    """
    Returns (category, reason) where category is one of:
    'Contracts', 'P&L', 'Unrelated'
    """
    if not GEMINI_API_KEY:
        print("  ⚠  No API key set — defaulting to 'Unrelated'")
        return "Unrelated", "No API key configured"

    url = (
        f"https://generativelanguage.googleapis.com/v1beta/"
        f"models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    )

    prompt = f"""You are a document classifier for DCT (Department of Culture and Tourism).
Classify the document below into EXACTLY ONE of the following categories:

1. "Contracts"  — Legal agreements, service contracts, vendor agreements, event contracts,
                  purchase orders, NDAs, MoUs, terms & conditions, sponsorship agreements,
                  letters of intent, facility-use agreements
2. "P&L"        — Profit & Loss statements, financial reports, income statements,
                  revenue & expense summaries, financial performance reports, budget actuals
3. "Unrelated"  — Everything else: invitations, presentations, event programs, floor plans,
                  schedules, menus, photos or image descriptions, brochures, general
                  correspondence, press releases, social-media content

Document filename : {filename}
Document text (excerpt):
{text[:MAX_TEXT_CHARS]}

Return ONLY valid JSON (no markdown, no explanation):
{{"category": "Contracts" | "P&L" | "Unrelated", "reason": "one concise sentence"}}"""

    payload = {"contents": [{"parts": [{"text": prompt}]}]}

    for attempt in range(retries + 1):
        try:
            resp = requests.post(url, json=payload, timeout=30)

            if resp.status_code == 200:
                raw   = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
                clean = re.sub(r"```json\s*|```", "", raw).strip()
                match = re.search(r"\{.*?\}", clean, re.DOTALL)
                if match:
                    result   = json.loads(match.group())
                    category = result.get("category", "Unrelated")
                    if category not in CATEGORIES:
                        category = "Unrelated"
                    return category, result.get("reason", "")
                raise ValueError("No JSON object in AI response")

            elif resp.status_code == 429 and attempt < retries:
                wait = 45
                print(f"  ⏳  Rate limit — waiting {wait}s (retry {attempt + 1}/{retries})...")
                time.sleep(wait)
                continue

            else:
                msg = f"API error {resp.status_code}: {resp.text[:200]}"
                print(f"  ⚠  {msg}")
                return "Unrelated", msg

        except Exception as e:
            print(f"  ⚠  Request failed: {e}")
            if attempt == retries:
                return "Unrelated", f"Error: {e}"
            time.sleep(3)

    return "Unrelated", "Max retries exceeded"


# ── STEP 6: Main classifier ────────────────────────────────────

def run_classifier():
    input_path = Path(INPUT_FOLDER)

    if not input_path.exists():
        print(f"❌  Input folder not found:\n   {INPUT_FOLDER}")
        print("   Update INPUT_FOLDER in Step 2 and re-run.")
        return

    if not GEMINI_API_KEY:
        print("❌  GEMINI_API_KEY is empty.")
        print("   Add it to Colab Secrets (key: GOOGLE_API_KEY) or paste it in Step 2.")
        return

    # ── Create output folder ──────────────────────────────────
    ts         = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = Path(OUTPUT_PARENT) / f"Output_{ts}"
    for cat in CATEGORIES:
        (output_dir / cat).mkdir(parents=True, exist_ok=True)
    print(f"📁  Output folder: {output_dir}\n")

    # ── Collect files recursively ─────────────────────────────
    all_files = sorted([
        p for p in input_path.rglob("*")
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS
    ])

    if not all_files:
        print(f"⚠   No supported files found in: {INPUT_FOLDER}")
        print(f"    Supported types: {', '.join(sorted(SUPPORTED_EXTS))}")
        return

    print(f"Found {len(all_files)} document(s) to classify.\n")
    col_w = min(max(len(p.name) for p in all_files) + 2, 50)
    header = f"{'#':<4} {'File':<{col_w}} {'Category':<12} Reason"
    print(header)
    print("─" * min(len(header) + 20, 110))

    summary = {cat: [] for cat in CATEGORIES}
    no_text = []

    for idx, file_path in enumerate(all_files, 1):
        short_name = file_path.name
        print(f"{idx:<4} {str(short_name):<{col_w}}", end="", flush=True)

        # Extract text
        text = extract_text(file_path)
        if not text.strip():
            category, reason = "Unrelated", "No extractable text — filed as Unrelated"
            no_text.append(file_path.name)
            print(f"{'Unrelated':<12} (no text — scanned image?)")
        else:
            category, reason = classify_document(file_path.name, text)
            print(f"{category:<12} {reason[:70]}")

        # Copy file, handling name collisions
        dest = output_dir / category / file_path.name
        if dest.exists():
            stem, suffix = file_path.stem, file_path.suffix
            counter = 2
            while dest.exists():
                dest = output_dir / category / f"{stem}_{counter}{suffix}"
                counter += 1
        shutil.copy2(file_path, dest)
        summary[category].append(str(file_path.relative_to(input_path)))

        # Rate-limit friendly delay between files
        if idx < len(all_files):
            time.sleep(DELAY_BETWEEN_FILES)

    # ── Print summary ─────────────────────────────────────────
    total = len(all_files)
    print("\n" + "=" * 60)
    print("  CLASSIFICATION SUMMARY")
    print("=" * 60)
    for cat in CATEGORIES:
        count = len(summary[cat])
        bar   = ("█" * count + "░" * (total - count)) if total <= 40 else f"{count}/{total}"
        print(f"  {cat:<12}: {count:>3} file(s)  {bar}")
    print(f"  {'─'*38}")
    print(f"  {'Total':<12}: {total:>3} file(s)")

    if no_text:
        print(f"\n  ⚠  {len(no_text)} file(s) had no extractable text (filed as Unrelated):")
        for n in no_text:
            print(f"      - {n}")

    print("=" * 60)
    print(f"\n✅  Done!  Output saved to:\n   {output_dir}\n")

    # Pretty listing of output structure
    print("Output structure:")
    for cat in CATEGORIES:
        items = list((output_dir / cat).iterdir())
        print(f"  {output_dir.name}/{cat}/  ({len(items)} file(s))")


# ── RUN ────────────────────────────────────────────────────────
run_classifier()
