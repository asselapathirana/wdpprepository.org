import sys
from pathlib import Path
from datetime import datetime
from git import Repo
import win32com.client as win32
import re
import pypandoc
import chardet

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_DIR = SCRIPT_DIR.parent
DOCS_DIR = REPO_DIR / "docs"
CSS_URL = "https://wdpprepository.org/static/css/project.css"
REMOTE = "origin"
BRANCH = "main"


def convert_docx_to_html(docx_path: Path, output_path: Path):
    """Export Word → HTML via COM, fix encoding, and inject site CSS."""
    print(f"Converting {docx_path.name} using Word COM...")

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(str(docx_path))

    try:
        doc.WebOptions.Encoding = win32.constants.msoEncodingUTF8
    except Exception:
        print("⚠️  Could not set WebOptions.Encoding; will fix after export.")

    tmp_html = str(output_path)
    doc.SaveAs2(tmp_html, FileFormat=win32.constants.wdFormatFilteredHTML)
    doc.Close(False)
    word.Quit()

    raw = Path(tmp_html).read_bytes()

    # Detect actual encoding
    det = chardet.detect(raw)
    enc = det.get("encoding") or "cp1252"
    print(f"Detected encoding: {enc}")

    html = raw.decode(enc, errors="replace")

    # If signs of double encoding
    if any(x in html for x in ["â", "Â", "€"]):
        try:
            html = html.encode("latin1").decode("utf-8")
            print("Fixed double-encoded characters (latin1→utf-8).")
        except Exception:
            pass

    # Clean up + inject CSS
    html = re.sub(r"<style[^>]*>.*?</style>", "", html, flags=re.DOTALL | re.IGNORECASE)
    html = re.sub(r'charset=[^">]+', 'charset=utf-8', html, flags=re.IGNORECASE)
    css_link = f'<link rel="stylesheet" href="{CSS_URL}">\n'
    html = re.sub(r"</head>", css_link + "</head>", html, count=1, flags=re.IGNORECASE)

    # now convert HTML → Markdown
    md_text = pypandoc.convert_text(html, "markdown", format="html")
    md_path =  REPO_DIR / Path(f"{docx_path.stem}.md")
    print(md_path)
    md_path.write_text(md_text, encoding="utf-8")
    
    html = pypandoc.convert_text(md_text, "html", format="markdown")
    Path(output_path).write_text(html, encoding="utf-8")
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

    DOCS_DIR.mkdir(exist_ok=True)
    output_html = DOCS_DIR / f"{word_file.stem}.html"

    convert_docx_to_html(word_file, output_html)
    commit_and_push(REPO_DIR, f"Update {word_file.name} → HTML ({datetime.now():%Y-%m-%d %H:%M})")

    print(f"🌐 https://asselapathirana.github.io/wdpprepository.org/{output_html.name}")


if __name__ == "__main__":
    main()
