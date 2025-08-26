import os
import pandas as pd
from pathlib import Path
import fitz  # PyMuPDF
import re
from PIL import Image
import io
import sys
import traceback

def test_basic_functionality():
    """Test basic PDF reading"""
    print("Testing basic functionality...")
    
    try:
        # Test PyMuPDF
        print("✓ PyMuPDF imported successfully")
        
        # Test pandas
        test_df = pd.DataFrame({'test': [1, 2, 3]})
        print("✓ Pandas working")
        
        # Test PIL
        test_img = Image.new('RGB', (100, 100), color='red')
        print("✓ PIL working")
        
        return True
    except Exception as e:
        print(f"✗ Basic test failed: {e}")
        return False

def extract_model_name_simple(pdf_path):
    """Simple model name extraction"""
    filename = Path(pdf_path).stem
    model_match = re.search(r'(WM\d+\w*)', filename, re.IGNORECASE)
    if model_match:
        return model_match.group(1).upper()
    return filename

def find_spare_parts_page_simple(doc):
    """Find spare parts page with detailed logging"""
    print("  Searching for 'SUGGESTED SPARE PARTS' page...")
    
    for page_num in range(len(doc)):
        try:
            page = doc[page_num]
            text = page.get_text()
            print(f"    Page {page_num + 1}: {len(text)} characters")
            
            if "SUGGESTED SPARE PARTS" in text.upper():
                print(f"    ✓ Found 'SUGGESTED SPARE PARTS' on page {page_num + 1}")
                return page_num
                
        except Exception as e:
            print(f"    ✗ Error reading page {page_num + 1}: {e}")
            continue
    
    print("    No 'SUGGESTED SPARE PARTS' page found")
    return None

def simple_table_check(text):
    """Simple check for table content"""
    table_indicators = ['NO.', 'PART NO.', 'PART NAME', 'QTY', 'REMARKS']
    count = sum(1 for indicator in table_indicators if indicator in text.upper())
    return count >= 2

def extract_spare_parts_basic(doc, page_num, model_name):
    """Basic spare parts extraction"""
    print(f"  Extracting spare parts from page {page_num + 1}")
    
    try:
        page = doc[page_num]
        text = page.get_text()
        lines = text.split('\n')
        
        spare_parts_data = []
        
        for line in lines:
            line = line.strip()
            if not line or len(line) < 10:
                continue
            
            # Simple pattern matching for spare parts
            # Pattern: digits...part_number...description
            if re.search(r'^\d+.*[A-Z0-9\-]{5,}.*[A-Za-z]', line):
                parts = re.split(r'\.{3,}|\s{3,}', line)
                if len(parts) >= 3:
                    clean_parts = []
                    for part in parts:
                        clean_part = part.strip(' .')
                        if clean_part:
                            clean_parts.append(clean_part)
                    
                    if len(clean_parts) >= 3:
                        spare_parts_data.append({
                            'Model': model_name,
                            'Quantity': clean_parts[0],
                            'Part Number': clean_parts[1] if len(clean_parts) > 1 else '',
                            'Description': ' '.join(clean_parts[2:]) if len(clean_parts) > 2 else ''
                        })
        
        print(f"    Extracted {len(spare_parts_data)} spare parts entries")
        return spare_parts_data
        
    except Exception as e:
        print(f"    Error extracting spare parts: {e}")
        return []

def simple_content_analysis(page):
    """Simple analysis of page content"""
    try:
        text = page.get_text()
        
        # Count characteristics
        images = len(page.get_images())
        drawings = len(page.get_drawings())
        text_length = len(text.strip())
        
        # Check for table indicators
        is_table = simple_table_check(text)
        
        print(f"    Content analysis: {images} images, {drawings} drawings, {text_length} chars, table={is_table}")
        
        return {
            'has_images': images > 0,
            'has_drawings': drawings > 10,
            'is_table': is_table,
            'text_length': text_length
        }
        
    except Exception as e:
        print(f"    Error analyzing content: {e}")
        return {'has_images': False, 'has_drawings': False, 'is_table': False, 'text_length': 0}

def extract_table_basic(doc, page_num, model_name):
    """Basic table extraction"""
    print(f"  Attempting to extract table from page {page_num + 1}")
    
    try:
        page = doc[page_num]
        text = page.get_text()
        
        analysis = simple_content_analysis(page)
        
        if not analysis['is_table']:
            print("    No table structure detected")
            return []
        
        print("    Table structure detected, extracting data...")
        
        lines = text.split('\n')
        table_rows = []
        page_title = ""
        
        # Get page title
        for line in lines[:10]:
            line = line.strip()
            if len(line) > 5 and not line.startswith('PAGE') and not line.startswith('WM63SLF'):
                page_title = line
                break
        
        # Extract table rows
        for line in lines:
            line = line.strip()
            if not line or len(line) < 10:
                continue
            
            # Skip headers and page info
            if any(word in line.upper() for word in ['NO.', 'PART NO.', 'PART NAME', 'QTY', 'REMARKS', 'PAGE']):
                continue
            
            if line.upper() == page_title.upper():
                continue
            
            # Look for data rows
            row_data = None
            
            # Pattern 1: "1    EM948630    DECAL, PUSH TO STOP    1"
            parts = re.split(r'\s{2,}', line)
            if len(parts) >= 3:
                clean_parts = [p.strip() for p in parts if p.strip()]
                if len(clean_parts) >= 3 and re.match(r'^\d+$', clean_parts[0]):
                    row_data = clean_parts
            
            # Pattern 2: Dot-separated
            if not row_data and '...' in line:
                parts = re.split(r'\.{3,}', line)
                if len(parts) >= 2:
                    clean_parts = [p.strip(' .') for p in parts if p.strip(' .')]
                    if len(clean_parts) >= 2:
                        row_data = clean_parts
            
            if row_data and len(row_data) >= 2:
                # Pad to 5 columns
                while len(row_data) < 5:
                    row_data.append("")
                table_rows.append(row_data[:5])
        
        if table_rows:
            table_data = {
                'page': page_num + 1,
                'model': model_name,
                'title': page_title or f"Table - Page {page_num + 1}",
                'rows': table_rows
            }
            print(f"    Extracted {len(table_rows)} table rows")
            return [table_data]
        
    except Exception as e:
        print(f"    Error extracting table: {e}")
        traceback.print_exc()
    
    return []

def extract_image_basic(page, page_num, model_name, output_folder):
    """Basic image extraction"""
    print(f"  Checking page {page_num + 1} for images...")
    
    try:
        analysis = simple_content_analysis(page)
        
        # Skip if it's clearly a table
        if analysis['is_table']:
            print("    Page is a table - skipping image extraction")
            return []
        
        images_folder = Path(output_folder) / "images" / model_name
        images_folder.mkdir(parents=True, exist_ok=True)
        
        extracted_images = []
        
        # Extract embedded images
        image_list = page.get_images()
        if image_list:
            print(f"    Found {len(image_list)} embedded images")
            
            for img_index, img in enumerate(image_list):
                try:
                    xref = img[0]
                    pix = fitz.Pixmap(page.parent, xref)
                    
                    if pix.width > 50 and pix.height > 50:
                        img_filename = f"{model_name}_page_{page_num + 1:02d}_img_{img_index + 1}.png"
                        img_path = images_folder / img_filename
                        pix.save(str(img_path))
                        extracted_images.append(str(img_path))
                        print(f"      Saved: {img_filename}")
                    
                    pix = None
                    
                except Exception as e:
                    print(f"      Error extracting image {img_index + 1}: {e}")
        
        # Save page as diagram if it has many drawings and little text
        if analysis['has_drawings'] and analysis['text_length'] < 1000:
            try:
                print(f"    Saving page as diagram ({analysis['text_length']} chars, many drawings)")
                mat = fitz.Matrix(2.0, 2.0)
                pix = page.get_pixmap(matrix=mat)
                diagram_filename = f"{model_name}_page_{page_num + 1:02d}_diagram.png"
                diagram_path = images_folder / diagram_filename
                pix.save(str(diagram_path))
                extracted_images.append(str(diagram_path))
                print(f"      Saved diagram: {diagram_filename}")
                pix = None
            except Exception as e:
                print(f"      Error saving diagram: {e}")
        
        return extracted_images
        
    except Exception as e:
        print(f"    Error processing images: {e}")
        return []

def process_pdf_simple(pdf_path, output_folder):
    """Simple PDF processing with extensive logging"""
    print(f"\n{'='*60}")
    print(f"PROCESSING: {pdf_path}")
    print(f"{'='*60}")
    
    if not Path(pdf_path).exists():
        print(f"ERROR: PDF file does not exist: {pdf_path}")
        return None
    
    model_name = extract_model_name_simple(pdf_path)
    print(f"Model: {model_name}")
    
    results = {
        'model_name': model_name,
        'spare_parts_data': [],
        'images': [],
        'tables': [],
        'status': 'unknown'
    }
    
    try:
        print("Opening PDF...")
        doc = fitz.open(pdf_path)
        print(f"PDF opened successfully - {len(doc)} pages")
        
        # Find spare parts page
        spare_parts_page = find_spare_parts_page_simple(doc)
        
        if spare_parts_page is not None:
            # Extract spare parts
            spare_parts_data = extract_spare_parts_basic(doc, spare_parts_page, model_name)
            results['spare_parts_data'] = spare_parts_data
            
            # Process pages after spare parts
            print(f"\nProcessing pages after spare parts (starting from page {spare_parts_page + 2})...")
            
            for page_num in range(spare_parts_page + 1, min(len(doc), spare_parts_page + 6)):  # Limit for testing
                print(f"\n--- Page {page_num + 1} ---")
                
                try:
                    page = doc[page_num]
                    
                    # Analyze content
                    analysis = simple_content_analysis(page)
                    
                    # Extract tables
                    if analysis['is_table']:
                        print("  Processing as TABLE")
                        tables = extract_table_basic(doc, page_num, model_name)
                        results['tables'].extend(tables)
                    else:
                        print("  Processing as potential IMAGE/DIAGRAM")
                        images = extract_image_basic(page, page_num, model_name, output_folder)
                        results['images'].extend(images)
                    
                except Exception as e:
                    print(f"  ERROR processing page {page_num + 1}: {e}")
                    traceback.print_exc()
                    continue
        
        doc.close()
        results['status'] = 'success'
        
        print(f"\nSUMMARY:")
        print(f"  Spare parts: {len(results['spare_parts_data'])}")
        print(f"  Tables: {len(results['tables'])}")
        print(f"  Images: {len(results['images'])}")
        
        return results
        
    except Exception as e:
        print(f"CRITICAL ERROR: {e}")
        traceback.print_exc()
        results['status'] = f'error: {str(e)}'
        return results

def save_results_simple(results, output_folder):
    """Simple results saving"""
    print(f"\nSaving results...")
    output_folder = Path(output_folder)
    
    files_created = []
    
    try:
        # Save spare parts
        if results['spare_parts_data']:
            df = pd.DataFrame(results['spare_parts_data'])
            spare_parts_path = output_folder / "spare_parts.xlsx"
            df.to_excel(spare_parts_path, index=False)
            files_created.append(spare_parts_path)
            print(f"✓ Spare parts saved: {spare_parts_path}")
        
        # Save tables
        if results['tables']:
            tables_path = output_folder / "tables.xlsx"
            
            with pd.ExcelWriter(tables_path, engine='openpyxl') as writer:
                for i, table in enumerate(results['tables']):
                    sheet_name = f"Page_{table['page']}"
                    if table.get('title'):
                        # Clean sheet name
                        clean_title = re.sub(r'[^\w\s]', '', table['title'])[:20]
                        sheet_name = f"P{table['page']}_{clean_title}"
                    
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:31]
                    
                    df = pd.DataFrame(table['rows'])
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            
            files_created.append(tables_path)
            print(f"✓ Tables saved: {tables_path}")
        
        return files_created
        
    except Exception as e:
        print(f"Error saving results: {e}")
        traceback.print_exc()
        return files_created

def main():
    """Main function with debug mode"""
    print("PDF PROCESSOR - DEBUG VERSION")
    print("=" * 50)
    
    # Test basic functionality
    if not test_basic_functionality():
        input("Press Enter to exit...")
        return
    
    # Setup
    PDF_FOLDER = "input_pdfs"
    OUTPUT_FOLDER = "output"
    
    Path(PDF_FOLDER).mkdir(exist_ok=True)
    Path(OUTPUT_FOLDER).mkdir(parents=True, exist_ok=True)
    
    # Find PDFs
    pdf_files = list(Path(PDF_FOLDER).glob("*.pdf"))
    
    if not pdf_files:
        print(f"No PDF files found in '{PDF_FOLDER}'")
        input("Press Enter to exit...")
        return
    
    print(f"Found {len(pdf_files)} PDF files:")
    for pdf in pdf_files:
        print(f"  - {pdf.name}")
    
    # Process first PDF only for debugging
    print(f"\nProcessing first PDF only (debug mode)...")
    
    result = process_pdf_simple(pdf_files[0], OUTPUT_FOLDER)
    
    if result:
        files_created = save_results_simple(result, OUTPUT_FOLDER)
        
        print(f"\nFiles created:")
        for file_path in files_created:
            print(f"  - {file_path}")
    
    print(f"\nDEBUG COMPLETE")
    input("Press Enter to exit...")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"FATAL ERROR: {e}")
        traceback.print_exc()
        input("Press Enter to exit...")