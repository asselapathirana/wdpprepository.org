import sys
import os
from pathlib import Path
import re
from datetime import datetime
import pypandoc
from git import Repo
from docx import Document

# run only first time pypandoc.download_pandoc()

# === Locate folders dynamically ===
SCRIPT_DIR = Path(__file__).resolve().parent
REPO_DIR = SCRIPT_DIR.parent
DOCS_DIR = REPO_DIR / "docs"
CSS_URL = "https://wdpprepository.org/static/css/project.css"
REMOTE = "origin"
BRANCH = "main"
# ==================================

def _debug_rewrite_image_paths(html: str, media_dir: Path, output_stem: str, debug: bool = True) -> str:
    """
    Rewrite absolute image paths pointing to media_dir → relative '<stem>_files/...'
    and print useful diagnostics.
    """
    # Build both Windows and POSIX-looking variants of the absolute media path
    media_abs = str(media_dir.resolve())
    media_win = media_abs.replace('/', '\\')
    media_posix = media_abs.replace('\\', '/')

    # Pattern that matches either prefix (with trailing slash or backslash)
    pat = re.compile(
        rf"{re.escape(media_win)}[\\/]|{re.escape(media_posix)}/",
        flags=re.IGNORECASE
    )

    # --- Debug: show context before rewrite ---
    if debug:
        print("\n[DEBUG] media_dir (resolved):")
        print("  win :", media_win)
        print("  posix:", media_posix)

        # Peek at first few <img> src attributes
        srcs = re.findall(r'<img[^>]+src=["\']([^"\']+)["\']', html, flags=re.IGNORECASE)
        print(f"[DEBUG] Found {len(srcs)} <img> tags.")
        for i, s in enumerate(srcs[:5], 1):
            print(f"  [img {i}] {s}")

    # Do the replacement
    new_html, n_subs = pat.subn(f"{output_stem}_files/", html)

    # --- Debug: after rewrite ---
    if debug:
        print(f"[DEBUG] Replacements made: {n_subs}")
        new_srcs = re.findall(r'<img[^>]+src=["\']([^"\']+)["\']', new_html, flags=re.IGNORECASE)
        for i, s in enumerate(new_srcs[:5], 1):
            print(f"  [img {i} after] {s}")

        # If nothing changed, show a small snippet around the first <img> for inspection
        if n_subs == 0 and srcs:
            snippet = re.search(r'<img[^>]+>', html, flags=re.IGNORECASE)
            if snippet:
                start = max(0, snippet.start() - 120)
                end = min(len(html), snippet.end() + 120)
                print("\n[DEBUG] First <img> snippet (no replacements happened):")
                print(html[start:end])
                print("[DEBUG] ^ Check if the src starts with the media_dir above.")

    return new_html



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

    # 🔧 Rewrite absolute paths to relative, with debug prints
    html = _debug_rewrite_image_paths(html, media_dir, output_path.stem, debug=True)

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
    repo.git.add(A=True)
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

    # --- print live URL ---
    base_url = "https://asselapathirana.github.io/wdpprepository.org"
    print(f"\n🌐 Live page URL:")
    print(f"{base_url}/{output_html.name}")
    print()


if __name__ == "__main__":
    main()
