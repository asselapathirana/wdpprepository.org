import sys
from pathlib import Path
from datetime import datetime
import pypandoc
from git import Repo
from docx import Document

pypandoc.download_pandoc()

# === Locate folders dynamically ===
SCRIPT_DIR = Path(__file__).resolve().parent
REPO_DIR = SCRIPT_DIR.parent
DOCS_DIR = REPO_DIR / "docs"
CSS_URL = "https://wdpprepository.org/static/css/project.css"
REMOTE = "origin"
BRANCH = "main"
# ==================================


def get_title_from_word(docx_path: Path) -> str:
    """Return first Heading 1 text, or filename if none."""
    try:
        doc = Document(docx_path)
        for p in doc.paragraphs:
            if p.style.name.lower().startswith("heading 1"):
                text = p.text.strip()
                if text:
                    return text
        return docx_path.stem
    except Exception as e:
        print(f"⚠️  Could not extract title: {e}")
        return docx_path.stem


def convert_docx_to_html(docx_path: Path, output_path: Path, title: str):
    media_dir = output_path.parent / f"{output_path.stem}_files"
    media_dir.mkdir(exist_ok=True)
    extra_args = ["--extract-media", str(media_dir)]

    print(f"Converting {docx_path.name} → HTML …")
    html = pypandoc.convert_file(str(docx_path), "html", extra_args=extra_args)

    wrapped = f"""---
layout: none
title: "{title}"
---
<html>
<head>
<meta charset="utf-8">
<link rel="stylesheet" href="{CSS_URL}">
<style>
  body {{ max-width: 900px; margin: 2rem auto; padding: 1rem; background: white; }}
  figure {{ text-align: center; margin: 2rem auto; }}
  figcaption {{ font-style: italic; color: #555; margin-top: 0.5rem; }}
</style>
</head>
<body>
{html}
</body>
</html>"""

    output_path.write_text(wrapped, encoding="utf-8")
    print(f"✅ HTML created: {output_path.relative_to(REPO_DIR)}")
    print(f"🖼️ Media saved in: {media_dir.relative_to(REPO_DIR)}")


def commit_and_push(repo_dir: Path, message: str):
    repo = Repo(repo_dir)
    repo.git.add(all=True)
    if repo.is_dirty():
        repo.index.commit(message)
        repo.git.push(REMOTE, BRANCH)
        print("🚀 Changes pushed to GitHub Pages.")
    else:
        print("No new changes detected.")


def main():
    if len(sys.argv) < 2:
        print("Usage: python authoring/publish_word_page.py <word_file>")
        sys.exit(1)

    # Handle relative paths gracefully (even if run from repo root)
    word_file = Path(sys.argv[1])
    if not word_file.is_absolute():
        word_file = (SCRIPT_DIR / word_file).resolve()

    if not word_file.exists():
        print(f"❌ Word file not found: {word_file}")
        sys.exit(1)

    title = get_title_from_word(word_file)
    DOCS_DIR.mkdir(exist_ok=True)
    output_html = DOCS_DIR / f"{word_file.stem}.html"

    convert_docx_to_html(word_file, output_html, title)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    commit_and_push(REPO_DIR, f"Update {word_file.name} → HTML ({timestamp})")


if __name__ == "__main__":
    main()
