import os
import sys
import fitz  # PyMuPDF
import comtypes.client
from pathlib import Path
from pdf2pptx import convert_pdf2pptx

# Global state
BULK_DELETE = False

def normalize_path_input(raw):
    raw = raw.strip().replace('"', '').replace("'", "")
    return Path(raw).expanduser()

def ask(prompt, default=None):
    res = input(prompt).strip()
    if res == "" and default is not None:
        return default
    return res

def pptx_to_pdf_clean(pptx_path: Path, pdf_path: Path):
    powerpoint = None
    try:
        # Re-initializing COM to ensure clean state
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        # MUST be visible to avoid the "Hiding window not allowed" error
        powerpoint.Visible = 1 
        
        deck = powerpoint.Presentations.Open(str(pptx_path.absolute()), WithWindow=False)
        deck.SaveAs(str(pdf_path.absolute()), 32)
        deck.Close()
    finally:
        if powerpoint:
            powerpoint.Quit()

def convert_file(file_path: Path, mode: str):
    out_ext = ".pptx" if mode == "1" else ".pdf"
    out_path = file_path.with_suffix(out_ext)

    try:
        if mode == "1":
            doc = fitz.open(str(file_path))
            num_pages = len(doc)
            doc.close()
            convert_pdf2pptx(str(file_path), str(out_path), 300, False, num_pages)
        else:
            pptx_to_pdf_clean(file_path, out_path)
        
        print(f"✅ Success: {out_path.name}")

        if BULK_DELETE:
            file_path.unlink()
            print(f"   🗑️ Deleted original.")
        else:
            # If not bulk, ask for each one
            choice = ask(f"Delete original {file_path.name}? (y/n): ", default="n").lower()
            if choice == 'y':
                file_path.unlink()
                print(f"   🗑️ Deleted original.")

    except Exception as e:
        print(f"❌ Error with {file_path.name}: {e}")

def main():
    global BULK_DELETE
    print("--- PDF/PPTX CONVERTER ---")
    
    mode = ask("Choose: 1) PDF -> PPTX  2) PPTX -> PDF: ")
    raw = ask("Enter path to file or folder: ")
    
    # Simplified Y/N choice
    bulk_choice = ask("Delete all originals automatically? (y/n): ", default="n").lower()
    BULK_DELETE = (bulk_choice == 'y')

    path = normalize_path_input(raw)
    if not path.exists():
        print("Path not found."); sys.exit(1)

    if path.is_dir():
        pattern = "*.pdf" if mode == "1" else "*.pptx"
        files = list(path.rglob(pattern))
        for f in files:
            convert_file(f, mode)
    else:
        convert_file(path, mode)

if __name__ == "__main__":
    main()