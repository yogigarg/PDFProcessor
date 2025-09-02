import os
import pandas as pd
from pathlib import Path
import fitz  # PyMuPDF
import re
import traceback

# ---------- TABLE EXTRACTION ----------
def extract_table_basic(doc, page_num, model_name):
    """Improved table extraction with 5 fixed columns"""
    print(f"  Extracting table data from page {page_num + 1}")
    
    try:
        page = doc[page_num]
        text = page.get_text("text")
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        
        table_rows = []
        page_title = ""
        
        # Try to find a title
        for line in lines[:10]:
            if len(line) > 5 and not line.upper().startswith("PAGE") and "WM63SLF" not in line:
                page_title = line
                break
        print(f"    Page title: {page_title}")
        
        buffer = []
        for line in lines:
            # Skip headers
            if any(h in line.upper() for h in ["NO.", "PART NO.", "PART NAME", "QTY", "REMARKS"]):
                continue
            
            if re.match(r"^\d+", line):  # new row starts with number
                if buffer:
                    row = parse_buffer(buffer)
                    if row: table_rows.append(row)
                buffer = [line]
            else:
                if buffer:
                    buffer.append(line)
        
        # Flush last
        if buffer:
            row = parse_buffer(buffer)
            if row: table_rows.append(row)
        
        if table_rows:
            print(f"    ✓ Extracted {len(table_rows)} rows")
            return [{
                "page": page_num + 1,
                "model": model_name,
                "title": page_title or f"Page {page_num+1} Table",
                "rows": table_rows
            }]
        else:
            print(f"    ✗ No rows captured")
            return []
    
    except Exception as e:
        print(f"    ✗ Error extracting table: {e}")
        traceback.print_exc()
        return []

def parse_buffer(buffer):
    """Helper to clean up a buffered row and return 5 fixed columns"""
    combined = " ".join(buffer)
    parts = re.split(r'\s{2,}|\.{3,}', combined)
    row = [p.strip(" .") for p in parts if p.strip(" .")]
    if not row:
        return None
    while len(row) < 5:
        row.append("")
    return row[:5]

# ---------- IMAGE EXTRACTION ----------
def extract_image_basic(page, page_num, model_name, output_folder):
    """Basic image/diagram extraction"""
    print(f"  Extracting images from page {page_num + 1}")
    try:
        images_folder = Path(output_folder) / "images" / model_name
        images_folder.mkdir(parents=True, exist_ok=True)
        extracted_images = []
        
        drawings = page.get_drawings()
        text_length = len(page.get_text().strip())
        
        # Save diagram if it has significant vector content
        if len(drawings) > 50:
            print(f"    Detected diagram with {len(drawings)} drawings, {text_length} chars")
            mat = fitz.Matrix(3.0, 3.0)
            pix = page.get_pixmap(matrix=mat)
            diagram_filename = f"{model_name}_page_{page_num + 1:02d}_diagram.png"
            diagram_path = images_folder / diagram_filename
            pix.save(str(diagram_path))
            extracted_images.append(str(diagram_path))
            print(f"      ✓ Saved diagram: {diagram_filename}")
        
        return extracted_images
    except Exception as e:
        print(f"    ✗ Error extracting image: {e}")
        traceback.print_exc()
        return []

# ---------- PDF PROCESSING ----------
def process_pdf(pdf_path, output_folder):
    """PDF processing with forced page classification"""
    print(f"\nPROCESSING: {pdf_path}")
    
    model_name = Path(pdf_path).stem.split("-")[0].upper()
    results = {"model_name": model_name, "tables": [], "images": []}
    
    doc = fitz.open(pdf_path)
    print(f"PDF opened successfully - {len(doc)} pages")
    
   # Forced classification: diagrams vs tables
    diagram_pages = {6, 8, 10, 12, 14, 16, 18, 20, 22, 24}

    if (page_num + 1) in diagram_pages:
   	 print("  → FORCED DIAGRAM PAGE")
   	 extracted_images = extract_image_basic(page, page_num, model_name, output_folder)
   	 results['images'].extend(extracted_images)
    else:
    	print("  → FORCED TABLE PAGE")
    	tables = extract_table_basic(doc, page_num, model_name)
    	results['tables'].extend(tables)
    
    doc.close()
    return results

# ---------- SAVE RESULTS ----------
def save_results_pagewise(results, output_folder):
    """Save tables with each page in a separate sheet"""
    output_folder = Path(output_folder)
    files_created = []
    
    if results["tables"]:
        tables_path = output_folder / "tables_by_page.xlsx"
        with pd.ExcelWriter(tables_path, engine="openpyxl") as writer:
            for table in results["tables"]:
                page_num = table["page"]
                title = table.get("title", f"Page{page_num}")
                clean_title = re.sub(r"[^\w\s]", "", title)
                clean_title = re.sub(r"\s+", "_", clean_title)
                sheet_name = f"P{page_num}_{clean_title}"[:31]
                
                df = pd.DataFrame(table["rows"], columns=["NO.", "PART NO.", "PART NAME", "QTY", "REMARKS"])
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"✓ Saved Page {page_num} → {sheet_name} ({len(df)} rows)")
        files_created.append(tables_path)
    
    return files_created

# ---------- MAIN ----------
if __name__ == "__main__":
    PDF_PATH = "WM63SLF-rev-0-parts-manual.pdf"
    OUTPUT_FOLDER = "output"
    Path(OUTPUT_FOLDER).mkdir(exist_ok=True)
    
    results = process_pdf(PDF_PATH, OUTPUT_FOLDER)
    save_results_pagewise(results, OUTPUT_FOLDER)
    print("\nProcessing complete.")
