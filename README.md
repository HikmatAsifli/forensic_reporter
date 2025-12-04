
# 🕵️‍♂️ Unified Digital Forensics Evidence Reporter

A Python tool for automatically extracting and documenting embedded or recovered images from PDF files and directories/ZIP archives into a structured forensic report (`.docx`). Designed for offensive security professionals, digital forensics analysts, and red teamers who need to catalog visual evidence with integrity hashes and metadata.


## Features

- **PDF Image Extraction:** Pulls all embedded images from every page of a PDF (using PyMuPDF).
- **Filesystem/ZIP Support:** Processes images from local directories or ZIP archives.
- **Forensic Metadata:**
    - SHA-256 hash of each image
    - File size
    - Filesystem timestamps (`mtime, ctime`)
    - Source context (PDF page/image index or file path)
- **Professional Report:** Generates a clean Microsoft Word (.docx) report with embedded thumbnails and metadata tables.
- **Cross-Platform:** Works on Kali Linux, Windows, macOS — anywhere Python runs.

🔍 Note: On Linux systems, “Created” time reflects `st_ctime` (inode change time), not true creation time.
## 🛠️ Requirements

- Python 3.7+
- Required packages:
```bash
pip install python-docx PyMuPDF
```

## 💡 Install all dependencies

```bash
pip install -r requirements.txt
```
## 📦 Usage

```bash
python forencis_reporter.py <pdf_file_or_NONE> <image_folder_or_zip> <output.docx>
```
## Arguments


| Argument       | Description                |
|  :------- | :------------------------- |
| `pdf_file_or_NONE` | Path to a PDF file, or `NONE` if not used |
| `image_folder_or_zip` | Path to a directory or ZIP file containing images, or `NONE` |
| `output.docx` | Output Word document filename |

## Examples

**1) PDF + Folder:**
```bash
python forencis_reporter.py malware_report.pdf extracted_images/ forensic_evidence.docx
```
**2) PDF only:**
```bash
python forencis_reporter.py screenshot_evidence.pdf NONE report.docx
```
**3) ZIP only:**
```bash
python forencis_reporter.py NONE exfil_pictures.zip incident_images.docx
```

⚠️ At least one input source (PDF or image folder/ZIP) must be provided.
## 🗂️ Supported Image Formats

- `.jpg` / `.jpeg`
- `.png`
- `.bmp`
- `.gif`
- `.tiff`
- `.webp`
Other formats may work if your system’s Word supports them, but are not guaranteed.
## 📄 Sample Report Structure

- Title: Unified Digital Forensics Evidence Report
- Timestamp: Generation time
- Section 1: Source: Embedded Images from PDF
    - Table per image: Thumbnail + Hash, Size, Source (e.g., “PDF Page 3, Image 2”)
- Section 2: Source: Recovered Image Files (Directory/ZIP)
    - Table per image: Thumbnail + Hash, Size, Modified/Created timestamps, File path
Each image is embedded directly into the document for easy review.
## 🧪 Use Cases

- Documenting visual artifacts from phishing PDFs
- Cataloging exfiltrated images during pentests
- Creating court-ready evidence logs
- Automating forensic triage in red team engagements
## ⚠️ Limitations

- Does not recover deleted/steganographic images — only processes visible or embedded ones.
- PDF image extraction depends on how the PDF was created (some vector/layered images may not extract as expected).
- Word report may become large with many high-res images.
## 📝 License
MIT License — feel free to use, modify, and integrate into your security toolchain.
## 📬 Contact

Coded with ☠️ by **Mishima** & **Nocturnis**

(Inspired by real-world offensive & forensic workflows)