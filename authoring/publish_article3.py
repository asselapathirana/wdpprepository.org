import sys
import os
from pathlib import Path
from datetime import datetime
from git import Repo
from docx import Document
import win32com.client as win32
import re

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
    """Use MS Word COM to export HTML; post-process for site layout."""
    print(f"Converting {docx_path.name} using Word…")
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False

    doc = word.Documents.Open(str(docx_path))
    html_tmp = str(output_path)
    doc.SaveAs(html_tmp, FileFormat=8)  # 8 = wdFormatHTML
    doc.Close(False)
    word.Quit()

    raw = Path(html_tmp).read_bytes()
    try:
        html = raw.decode("utf-8")
    except UnicodeDecodeError:
        html = raw.decode("cp1252")  # Word’s default for .htm export


    # remove MS inline styles and insert your CSS link
    html = re.sub(r"<style[^>]*>.*?</style>", "", html, flags=re.DOTALL | re.IGNORECASE)
    css_link = f'<link rel="stylesheet" href="{CSS_URL}">\n'
    html = re.sub(r"</head>", css_link + "</head>", html, count=1, flags=re.IGNORECASE)

    if not re.search(r"<h1\b", html, re.IGNORECASE):
        html = re.sub(r"<body([^>]*)>", rf"<body\1>\n<h1>{title}</h1>", html, 1, re.IGNORECASE)

    front_matter = f'---\nlayout: none\ntitle: "{title}"\n---\n'
    Path(output_path).write_text(front_matter + html, encoding="utf-8")

    print(f"✅ HTML saved: {output_path.relative_to(REPO_DIR)}")


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
