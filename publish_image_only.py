import sys
import shutil
import re
import subprocess
from pathlib import Path
from docx import Document

repo_url = "https://asselapathirana.github.io/wdpprepository.org"

def extract_images_with_captions_and_push(docx_path: Path):
    """Extract images and captions from a Word file, generate simple HTML pages, and push to GitHub."""
    doc = Document(docx_path)
    output_root = Path("docs") / docx_path.stem
    img_dir = output_root / "images"
    
    # clean old export
    if output_root.exists():
        shutil.rmtree(output_root)
    
    
    output_root.mkdir(parents=True, exist_ok=True)
    img_dir.mkdir(exist_ok=True)

    images = []
    img_count = 0
    pending_image = None  # track if last element was image

    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        style = p.style.name.lower() if p.style else ""

        # detect image(s) in this paragraph
        runs_with_images = [
            run for run in p.runs
            if run.element.xpath(".//pic:pic")
        ]
        if runs_with_images:
            for run in runs_with_images:
                for shape in run.element.xpath(".//pic:pic"):
                    img_count += 1
                    blip = shape.xpath(".//a:blip")[0]
                    rId = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                    part = run.part.related_parts[rId]
                    image_bytes = part.blob
                    ext = part.content_type.split("/")[-1]
                    filename = f"img_{img_count:03d}.{'jpg' if ext == 'jpeg' else ext}"
                    img_path = img_dir / filename
                    img_path.write_bytes(image_bytes)
                    pending_image = filename  # mark image awaiting caption

        # detect caption — explicit or implicit
        if pending_image:
            is_explicit_caption = style.startswith("caption")
            has_text = bool(text)
            if is_explicit_caption or (has_text and not runs_with_images):
                caption = text if text else f"Image {img_count}"
                safe_caption = re.sub(r"[^a-zA-Z0-9_-]+", "_", caption[:40]) or f"img_{img_count:03d}"
                new_filename = f"{Path(pending_image).stem}_{safe_caption}{Path(pending_image).suffix}"
                (img_dir / pending_image).rename(img_dir / new_filename)
                images.append((new_filename, caption))
                pending_image = None
            elif not text and not style.startswith("caption"):
                # Wait for next paragraph to possibly carry caption
                continue

    # handle leftover images without captions
    if pending_image:
        caption = f"Image {img_count}"
        images.append((pending_image, caption))

    # Create per-image pages
    for filename, caption in images:
        page_path = output_root / f"{Path(filename).stem}.html"
        html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>{caption}</title></head>
<body style="font-family:sans-serif;max-width:700px;margin:auto;text-align:center;">
  <img src="images/{filename}" style="max-width:100%;height:auto;"><br>
  <p>{caption}</p>
</body></html>"""
        page_path.write_text(html, encoding="utf-8")

    # Create index
    index_html = "<h1>Image Index</h1>\n<ul>\n"
    for filename, caption in images:
        stem = Path(filename).stem
        index_html += f'  <li><a href="{stem}.html">{caption}</a></li>\n'
    index_html += "</ul>\n"
    (output_root / "index.html").write_text(index_html, encoding="utf-8")

    print(f"✅ Extracted {len(images)} images to {output_root}")

    # Git commit and push
    try:
        subprocess.run(["git", "add", "-A", str(output_root)], check=True)
        subprocess.run(["git", "commit", "-m", f"Update images from {docx_path.name}"], check=True)
        subprocess.run(["git", "push"], check=True)
        print("🚀 Changes pushed to GitHub.")
        link = f"{repo_url}/docs/{docx_path.stem}/index.html"
        print(f"🌐 View it at: {link}")
    except subprocess.CalledProcessError as e:
        print(f"⚠️ Git push failed: {e}")

def main():
    if len(sys.argv) < 2:
        print("Usage: python extract_images.py <word_file.docx> (./authoring folder)")
        sys.exit(1)
    docx_path = Path('./authoring/'+sys.argv[1])
    if not docx_path.exists():
        print(f"❌ File not found: {docx_path}")
        sys.exit(1)
    extract_images_with_captions_and_push(docx_path)

if __name__ == "__main__":
    main()
