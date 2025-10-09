import sys
import os
from pathlib import Path
from datetime import datetime
from git import Repo
from docx import Document
import win32com.client as win32
import re
import win32com.client as win32

import chardet



SCRIPT_DIR = Path(__file__).resolve().parent
REPO_DIR = SCRIPT_DIR.parent
DOCS_DIR = REPO_DIR / "docs"
CSS_URL = "https://wdpprepository.org/static/css/project.css"
REMOTE = "origin"
BRANCH = "main"


def get_title_from_word(docx_path: Path) -> str:
    """Return title from Word metadata or first Heading 1."""
    try:
        doc = Document(docx_path)
        title_prop = (doc.core_properties.title or "").strip()
        if title_prop:
            print(f"✅ Found title: {title_prop}")
            return title_prop

        for p in doc.paragraphs:
            if getattr(p.style, "name", "").lower().startswith("heading 1"):
                if p.text.strip():
                    return p.text.strip()
        return docx_path.stem
    except Exception as e:
        print(f"❌ Title extraction failed: {e}")
        return docx_path.stem



def convert_docx_to_html(docx_path: Path, output_path: Path, title: str):
    """Export Word → HTML via COM, fix encoding, and inject site CSS."""
    print(f"Converting {docx_path.name} using Word COM...")

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(str(docx_path))

    # Try to enforce UTF-8 output
    try:
        doc.WebOptions.Encoding = win32.constants.msoEncodingUTF8
    except Exception:
        print("⚠️  Could not set WebOptions.Encoding; will fix after export.")

    tmp_html = str(output_path)
    doc.SaveAs2(tmp_html, FileFormat=win32.constants.wdFormatFilteredHTML)
    doc.Close(False)
    word.Quit()

    raw = Path(tmp_html).read_bytes()

    # --- Detect actual encoding ---
    det = chardet.detect(raw)
    enc = det.get("encoding") or "cp1252"
    confidence = det.get("confidence", 0)
    print(f"Detected encoding: {enc} (confidence {confidence:.2f})")

    # --- Decode safely ---
    html = raw.decode(enc, errors="replace")

    # If Word lied and we see garbage like â€™, fix by reinterpretation
    if any(x in html for x in ["â", "Â", "€"]):
        try:
            html = html.encode("latin1").decode("utf-8")
            print("Fixed double-encoded characters (latin1→utf-8).")
        except Exception:
            pass

    # --- Clean up ---
    html = re.sub(r"<style[^>]*>.*?</style>", "", html, flags=re.DOTALL | re.IGNORECASE)
    html = re.sub(r'charset=[^">]+', 'charset=utf-8', html, flags=re.IGNORECASE)
    css_link = f'<link rel="stylesheet" href="{CSS_URL}">\n'
    html = re.sub(r"</head>", css_link + "</head>", html, count=1, flags=re.IGNORECASE)

    if not re.search(r"<h1\b", html, re.IGNORECASE) and title:
        html = re.sub(r"<body([^>]*)>", rf"<body\1>\n<h1>{title}</h1>", html, 1, re.IGNORECASE)

    front_matter = f'---\nlayout: none\ntitle: "{title}"\n---\n'
    Path(output_path).write_text(front_matter + html, encoding="utf-8")

    print(f"✅ HTML saved: {output_path.name}")


def commit_and_push(repo_dir: Path, message: str):
    repo = Repo(repo_dir)
    repo.git.add(A=True)
    if repo.is_dirty():
        repo.index.commit(message)
        repo.git.push(REMOTE, BRANCH)
        print("🚀 Changes pushed.")
    else:
        print("No new changes detected.")


def main():
    if len(sys.argv) < 2:
        print("Usage: python publish_article3.py <word_file>")
        sys.exit(1)

    word_file = Path(sys.argv[1]).resolve()
    if not word_file.exists():
        print(f"❌ Not found: {word_file}")
        sys.exit(1)

    title = get_title_from_word(word_file)
    output_html = DOCS_DIR / f"{word_file.stem}.html"
    DOCS_DIR.mkdir(exist_ok=True)

    convert_docx_to_html(word_file, output_html, title)
    commit_and_push(REPO_DIR, f"Update {word_file.name} → HTML ({datetime.now():%Y-%m-%d %H:%M})")

    print(f"🌐 https://asselapathirana.github.io/wdpprepository.org/{output_html.name}")


if __name__ == "__main__":
    main()
