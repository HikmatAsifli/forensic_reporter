import os
import sys
import hashlib
import zipfile
import tempfile
import datetime
from pathlib import Path
from docx import Document
from docx.shared import Inches
import fitz  # PyMuPDF

# ANSI escape codes for color
RED = "\033[91m"
GREEN = "\033[92m"
YELLOW = "\033[93m"
BLUE = "\033[94m"
MAGENTA = "\033[95m"
CYAN = "\033[96m"
WHITE = "\033[97m"
RESET = "\033[0m"

print(f'''
{RED}      
                                                                                                                                               
@@@@@@@@  @@@@@@  @@@@@@@  @@@@@@@@ @@@  @@@  @@@@@@ @@@  @@@@@@@       @@@@@@@  @@@@@@@@ @@@@@@@   @@@@@@  @@@@@@@  @@@@@@@ @@@@@@@@ @@@@@@@  
@@!      @@!  @@@ @@!  @@@ @@!      @@!@!@@@ !@@     @@! !@@            @@!  @@@ @@!      @@!  @@@ @@!  @@@ @@!  @@@   @!!   @@!      @@!  @@@ 
@!!!:!   @!@  !@! @!@!!@!  @!!!:!   @!@@!!@!  !@@!!  !!@ !@!            @!@!!@!  @!!!:!   @!@@!@!  @!@  !@! @!@!!@!    @!!   @!!!:!   @!@!!@!  
!!:      !!:  !!! !!: :!!  !!:      !!:  !!!     !:! !!: :!!            !!: :!!  !!:      !!:      !!:  !!! !!: :!!    !!:   !!:      !!: :!!  
 :        : :. :   :   : : : :: ::  ::    :  ::.: :  :    :: :: :        :   : : : :: ::   :        : :. :   :   : :    :    : :: ::   :   : : 
                                                                                                                                               
{YELLOW}|-----------------------------------------------------{MAGENTA}Coded by Mishima and Nocturnis{YELLOW}----------------------------------------------------------------|{RESET}''')

# ------------------ Helpers ------------------

def sha256_data(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()

def sha256_file(filepath: str) -> str:
    hash_sha256 = hashlib.sha256()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_sha256.update(chunk)
    return hash_sha256.hexdigest()

def format_ts(ts: float) -> str:
    return datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')

def add_image_to_doc(doc: Document, image_path_or_bytes, metadata: dict):
    """Add image + metadata table to doc."""
    is_bytes = isinstance(image_path_or_bytes, bytes)
    ext = metadata.get("ext", "jpg")
    
    # Handle bytes by writing to temp file
    temp_path = None
    if is_bytes:
        # Use hash to avoid collisions
        hash_prefix = hashlib.md5(image_path_or_bytes).hexdigest()[:8]
        temp_path = f"temp_img_{hash_prefix}.{ext}"
        with open(temp_path, "wb") as f:
            f.write(image_path_or_bytes)
        display_path = temp_path
    else:
        display_path = image_path_or_bytes

    # Create table: image | metadata
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'

    # Embed image
    try:
        paragraph = table.cell(0, 0).paragraphs[0]
        run = paragraph.add_run()
        run.add_picture(display_path, width=Inches(3))
    except Exception as e:
        table.cell(0, 0).text = f"[Embed failed: {str(e)[:100]}]"

    # Build metadata string
    meta_text = f"SHA-256: {metadata['hash']}\n"
    if 'size' in metadata:
        meta_text += f"Size: {metadata['size']} bytes\n"
    if 'modified' in metadata:
        meta_text += f"Modified: {metadata['modified']}\n"
    if 'created' in metadata:
        meta_text += f"Created*: {metadata['created']}\n"
    meta_text += f"Source: {metadata['source']}"

    table.cell(0, 1).text = meta_text

    # Add warning if needed
    if metadata.get("warning"):
        table.cell(0, 1).add_paragraph(f"\n{metadata['warning']}")

    doc.add_paragraph()  # spacing

    # Cleanup temp file
    if temp_path and os.path.exists(temp_path):
        os.remove(temp_path)

# ------------------ Process PDF ------------------

def process_pdf(pdf_path: str, doc: Document):
    if not os.path.isfile(pdf_path):
        doc.add_paragraph(f"[!] PDF not found: {pdf_path}")
        return

    doc.add_heading("Source: Embedded Images from PDF", level=1)
    pdf_stat = os.stat(pdf_path)
    doc.add_paragraph(
        f"PDF Path: {os.path.abspath(pdf_path)}\n"
        f"PDF Modified: {format_ts(pdf_stat.st_mtime)}\n"
        f"Note: Embedded images have no original filesystem timestamps.\n"
    )

    pdf_doc = fitz.open(pdf_path)
    count = 0
    for page_num, page in enumerate(pdf_doc):
        for img_idx, img in enumerate(page.get_images(full=True)):
            count += 1
            xref = img[0]
            base_img = pdf_doc.extract_image(xref)
            img_bytes = base_img["image"]
            ext = base_img["ext"] or "jpg"

            img_hash = sha256_data(img_bytes)

            metadata = {
                "hash": img_hash,
                "ext": ext,
                "size": len(img_bytes),
                "source": f"PDF Page {page_num + 1}, Image {img_idx + 1}"
            }

            add_image_to_doc(doc, img_bytes, metadata)

    pdf_doc.close()
    doc.add_paragraph(f"Extracted {count} images from PDF.\n")

# ------------------ Process Folder or ZIP ------------------

def process_image_files(input_path: str, doc: Document):
    supported_ext = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp'}
    warning = "(*'Created' may reflect inode change time on Linux, not true NTFS creation time.)"

    if os.path.isdir(input_path):
        doc.add_heading("Source: Recovered Image Files (Directory)", level=1)
        image_files = sorted([f for f in Path(input_path).rglob('*') if f.suffix.lower() in supported_ext])
        for img in image_files:
            process_single_image_file(str(img), doc, warning)

    elif zipfile.is_zipfile(input_path):
        doc.add_heading("Source: Recovered Image Files (ZIP Archive)", level=1)
        with tempfile.TemporaryDirectory() as tmp:
            with zipfile.ZipFile(input_path, 'r') as zf:
                members = [m for m in zf.namelist() if Path(m).suffix.lower() in supported_ext and not m.endswith('/')]
                zf.extractall(tmp, members=members)
                for member in members:
                    full_path = os.path.join(tmp, member)
                    if os.path.isfile(full_path):
                        source_label = f"{input_path} (member: {member})"
                        process_single_image_file(full_path, doc, warning, source_label)
    else:
        doc.add_paragraph(f"[!] Not a valid directory or ZIP: {input_path}")

def process_single_image_file(img_path: str, doc: Document, warning: str = "", source_label: str = ""):
    try:
        stat = os.stat(img_path)
        metadata = {
            "hash": sha256_file(img_path),
            "size": stat.st_size,
            "modified": format_ts(stat.st_mtime),
            "created": format_ts(stat.st_ctime),
            "source": source_label if source_label else img_path,
            "warning": warning
        }
        add_image_to_doc(doc, img_path, metadata)
    except Exception as e:
        doc.add_paragraph(f"[!] Failed to process {img_path}: {e}\n")

# ------------------ Main ------------------

def main(pdf_path: str, file_source: str, output_docx: str):
    doc = Document()
    doc.add_heading('Unified Digital Forensics Evidence Report', 0)
    doc.add_paragraph(f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    if pdf_path and os.path.isfile(pdf_path):
        process_pdf(pdf_path, doc)
    else:
        doc.add_paragraph("No valid PDF provided.\n")

    if file_source and (os.path.isdir(file_source) or zipfile.is_zipfile(file_source)):
        process_image_files(file_source, doc)
    else:
        doc.add_paragraph("No valid image folder/ZIP provided.\n")

    doc.save(output_docx)
    print(f"[+] Report saved: {output_docx}")

# ------------------ CLI ------------------

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python forencis_reporter.py <pdf_file_or_NONE> <image_folder_or_zip> <output.docx>")
        print("Examples:")
        print("  python forencis_reporter.py evidence.pdf recovered/ report.docx")
        print("  python forencis_reporter.py NONE images.zip report.docx")
        print("  python forencis_reporter.py scan.pdf NONE report.docx")
        sys.exit(1)

    arg_pdf = sys.argv[1]
    arg_files = sys.argv[2]
    output = sys.argv[3]

    pdf_path = None if arg_pdf.lower() == "none" else arg_pdf
    file_source = None if arg_files.lower() == "none" else arg_files

    if not pdf_path and not file_source:
        print("[!] Error: Provide at least a PDF or image source.")
        sys.exit(1)

    main(pdf_path, file_source, output)