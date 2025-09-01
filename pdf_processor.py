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
    """Simple check for table content with debug info"""
    table_indicators = ['NO.', 'PART NO.', 'PART NAME', 'QTY', 'REMARKS']
    count = sum(1 for indicator in table_indicators if indicator in text.upper())
    
    # Also check for actual data rows that look like parts tables
    lines = text.split('\n')
    data_row_count = 0
    
    # Debug: show first few lines
    print(f"    Checking lines for table data:")
    for i, line in enumerate(lines[:15]):
        line = line.strip()
        if not line:
            continue
        print(f"      Line {i+1}: {line[:80]}{'...' if len(line) > 80 else ''}")
        
        # Look for lines that start with number and have part numbers
        if re.match(r'^\d+\s+[A-Z0-9\-]{4,}', line) or \
           re.match(r'^\d+\.+[A-Z0-9\-]{4,}', line) or \
           re.match(r'^\d+.*[A-Z0-9\-]{5,}.*[A-Za-z]', line):
            data_row_count += 1
            print(f"        ^ This looks like data row #{data_row_count}")
    
    # It's a table if we have headers OR multiple data rows
    is_table = count >= 2 or data_row_count >= 3
    print(f"    Table check: {count} headers, {data_row_count} data rows -> table={is_table}")
    return is_table

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
    """Basic table extraction - simplified"""
    print(f"  Extracting table data from page {page_num + 1}")
    
    try:
        page = doc[page_num]
        text = page.get_text()
        lines = text.split('\n')
        
        table_rows = []
        page_title = ""
        
        # Get page title
        for line in lines[:10]:
            line = line.strip()
            if len(line) > 5 and not line.startswith('PAGE') and not line.startswith('WM63SLF'):
                page_title = line
                break
        
        print(f"    Page title: {page_title}")
        
        # Extract table rows - be more aggressive about finding data
        for line in lines:
            line = line.strip()
            if not line or len(line) < 10:
                continue
            
            # Skip obvious headers and page info
            skip_line = any(word in line.upper() for word in [
                'NO.', 'PART NO.', 'PART NAME', 'QTY', 'REMARKS', 'PAGE', 'MIXER', 'MANUAL', 'REV.'
            ])
            
            if skip_line or line.upper() == page_title.upper():
                continue
            
            row_data = None
            
            # Pattern 1: "1    EM948630    DECAL, PUSH TO STOP    1"
            if re.match(r'^\d+\s+[A-Z0-9\-]{4,}', line):
                parts = re.split(r'\s{2,}', line)
                if len(parts) >= 3:
                    row_data = [p.strip() for p in parts if p.strip()]
                    print(f"      Found spaced row: {row_data}")
            
            # Pattern 2: "6............07055-034 .............................V-BELT, 4L340"
            elif '...' in line and re.match(r'^\d+\.+[A-Z0-9\-]{4,}', line):
                parts = re.split(r'\.{3,}', line)
                row_data = [p.strip(' .') for p in parts if p.strip(' .')]
                print(f"      Found dotted row: {row_data}")
            
            # Pattern 3: Any line that starts with digit and has part number pattern
            elif re.match(r'^\d+', line) and re.search(r'[A-Z0-9\-]{4,}', line):
                # Try different splitting approaches
                parts = re.split(r'\.{2,}|\s{3,}', line)
                if len(parts) >= 2:
                    row_data = [p.strip(' .') for p in parts if p.strip(' .')]
                    print(f"      Found general row: {row_data}")
            
            if row_data and len(row_data) >= 2:
                # Pad to 5 columns
                while len(row_data) < 5:
                    row_data.append("")
                table_rows.append(row_data[:5])
        
        if table_rows:
            table_data = {
                'page': page_num + 1,
                'model': model_name,
                'title': page_title or f"Parts Table - Page {page_num + 1}",
                'rows': table_rows
            }
            print(f"    ✓ Extracted {len(table_rows)} table rows")
            return [table_data]
        else:
            print(f"    ✗ No table rows found")
        
    except Exception as e:
        print(f"    ✗ Error extracting table: {e}")
        traceback.print_exc()
    
    return []

def extract_image_basic(page, page_num, model_name, output_folder):
    """Basic image extraction - simplified"""
    print(f"  Extracting images from page {page_num + 1}")
    
    try:
        images_folder = Path(output_folder) / "images" / model_name
        images_folder.mkdir(parents=True, exist_ok=True)
        
        extracted_images = []
        
        # Get page characteristics
        image_list = page.get_images()
        drawings = page.get_drawings()
        text = page.get_text()
        text_length = len(text.strip())
        
        # Extract embedded images first
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
                        print(f"      ✓ Saved embedded image: {img_filename}")
                    else:
                        print(f"      Skipped small image: {pix.width}x{pix.height}")
                    
                    pix = None
                    
                except Exception as e:
                    print(f"      ✗ Error extracting image {img_index + 1}: {e}")
        
        # Save as diagram if it has significant vector content (relaxed criteria)
        if len(drawings) > 100:  # Much lower threshold - if it has lots of drawings, it's likely a diagram
            print(f"    Detected technical diagram: {len(drawings)} drawings, {text_length} chars")
            try:
                mat = fitz.Matrix(3.0, 3.0)  # High resolution for diagrams
                pix = page.get_pixmap(matrix=mat)
                diagram_filename = f"{model_name}_page_{page_num + 1:02d}_diagram.png"
                diagram_path = images_folder / diagram_filename
                pix.save(str(diagram_path))
                extracted_images.append(str(diagram_path))
                print(f"      ✓ Saved technical diagram: {diagram_filename}")
                pix = None
            except Exception as e:
                print(f"      ✗ Error saving diagram: {e}")
        elif len(drawings) > 50:
            print(f"    Page has {len(drawings)} drawings - might be diagram but below threshold")
        else:
            print(f"    Page has only {len(drawings)} drawings - not a diagram")
        
        if not extracted_images:
            print(f"    No images extracted from this page")
        
        return extracted_images
        
    except Exception as e:
        print(f"    ✗ Error processing images: {e}")
        traceback.print_exc()
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
            
            # Process ALL pages after spare parts (removed limit)
            print(f"\nProcessing pages after spare parts (starting from page {spare_parts_page + 2})...")
            
            for page_num in range(spare_parts_page + 1, len(doc)):  # Process ALL remaining pages
                print(f"\n--- Page {page_num + 1} ---")
                
                try:
                    page = doc[page_num]
                    text = page.get_text()
                    
                    # Skip very short pages (likely just notes or page numbers)
                    if len(text.strip()) < 100:
                        print(f"  → SKIPPING (too little content: {len(text.strip())} chars)")
                        continue
                    
                    # Check content type ONCE
                    is_table = simple_table_check(text)
                    
                    if is_table:
                        print("  → PROCESSING AS TABLE")
                        tables = extract_table_basic(doc, page_num, model_name)
                        results['tables'].extend(tables)
                    else:
                        print("  → PROCESSING AS IMAGE/DIAGRAM")
                        # Show content analysis for image pages
                        images = len(page.get_images())
                        drawings = len(page.get_drawings())
                        text_length = len(text.strip())
                        print(f"    Content: {images} images, {drawings} drawings, {text_length} chars")
                        
                        extracted_images = extract_image_basic(page, page_num, model_name, output_folder)
                        results['images'].extend(extracted_images)
                    
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
    """Simple results saving with separate sheets for each table"""
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
        
        # Save tables with each table in a separate sheet
        if results['tables']:
            tables_path = output_folder / "tables.xlsx"
            
            print(f"  Creating Excel file with {len(results['tables'])} separate sheets...")
            
            with pd.ExcelWriter(tables_path, engine='openpyxl') as writer:
                for i, table in enumerate(results['tables']):
                    # Create meaningful sheet name
                    page_num = table['page']
                    title = table.get('title', '').strip()
                    model = table.get('model', 'Unknown')
                    
                    # Clean up title for sheet name
                    if title:
                        # Remove special characters and limit length
                        clean_title = re.sub(r'[^\w\s]', '', title)
                        clean_title = re.sub(r'\s+', '_', clean_title)
                        clean_title = clean_title[:20]  # Limit length
                        sheet_name = f"{model}_P{page_num}_{clean_title}"
                    else:
                        sheet_name = f"{model}_Page_{page_num}"
                    
                    # Ensure sheet name is unique and valid
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:31]
                    
                    print(f"    Sheet {i+1}: '{sheet_name}' ({len(table['rows'])} rows)")
                    
                    # Create DataFrame from table rows
                    if table['rows']:
                        # Add headers if available, otherwise use generic column names
                        max_cols = max(len(row) for row in table['rows']) if table['rows'] else 5
                        
                        if table.get('headers') and len(table['headers']) > 0:
                            # Use provided headers
                            headers = table['headers'][:max_cols]
                            # Pad headers if needed
                            while len(headers) < max_cols:
                                headers.append(f"Column_{len(headers)+1}")
                            df = pd.DataFrame(table['rows'], columns=headers)
                        else:
                            # Use generic headers
                            headers = ['NO.', 'PART_NO', 'PART_NAME', 'QTY', 'REMARKS'][:max_cols]
                            while len(headers) < max_cols:
                                headers.append(f"Column_{len(headers)+1}")
                            df = pd.DataFrame(table['rows'], columns=headers)
                        
                        # Save to Excel sheet
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # Add some formatting
                        worksheet = writer.sheets[sheet_name]
                        
                        # Auto-adjust column widths
                        for column in worksheet.columns:
                            max_length = 0
                            column_letter = column[0].column_letter
                            for cell in column:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except:
                                    pass
                            adjusted_width = min(max_length + 2, 50)
                            worksheet.column_dimensions[column_letter].width = adjusted_width
            
            files_created.append(tables_path)
            print(f"✓ Tables saved with separate sheets: {tables_path}")
        
        # Save processing summary
        summary_data = [{
            'Model': results['model_name'],
            'Status': results['status'],
            'Spare_Parts_Count': len(results['spare_parts_data']),
            'Tables_Count': len(results['tables']),
            'Images_Count': len(results['images']),
            'Table_Details': ', '.join([f"Page {t['page']}: {t.get('title', 'Untitled')}" for t in results['tables']])
        }]
        
        summary_df = pd.DataFrame(summary_data)
        summary_path = output_folder / "processing_summary.xlsx"
        summary_df.to_excel(summary_path, index=False)
        files_created.append(summary_path)
        print(f"✓ Processing summary saved: {summary_path}")
        
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