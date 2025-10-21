#!/usr/bin/env python3
"""
PDF Table Extractor Tool
Extracts table data from PDF files and outputs as JSON.
Looks for tables with Item, Description, Quantity, and UnitPrice columns.
"""

import argparse
import json
import re
import sys
from pathlib import Path
from typing import List, Dict, Any

try:
    import fitz  # PyMuPDF
except ImportError:
    print("Error: PyMuPDF is not installed. Please install it with: pip install PyMuPDF")
    sys.exit(1)


def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract all text from PDF file."""
    try:
        doc = fitz.open(pdf_path)
        full_text = ""
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            full_text += page.get_text()
        
        doc.close()
        return full_text
    except Exception as e:
        print(f"Error reading PDF file: {e}")
        sys.exit(1)


def parse_table_data(text: str) -> List[Dict[str, Any]]:
    """
    Parse table data from extracted text.
    Looks for items with A+6digit codes or 4digit codes and extracts Item, Description, Quantity, UnitPrice.
    Handles multi-line descriptions using trailing space detection.
    Goes through the entire document looking for item codes (A+6digits or 4digits).
    """
    items = []
    
    # Split text into lines but keep original lines for space checking
    all_lines = text.split('\n')
    lines = [line.strip() for line in all_lines if line.strip()]  # Cleaned lines for processing
    # Create a mapping from cleaned line index to original line
    original_line_map = {}
    cleaned_idx = 0
    for orig_idx, orig_line in enumerate(all_lines):
        if orig_line.strip():  # Only if line has content after stripping
            original_line_map[cleaned_idx] = orig_idx
            cleaned_idx += 1
    
    # Process the entire document, looking for A+6digit codes
    i = 0
    item_counter = 1
    while i < len(lines):
        try:
            line_content = lines[i].strip()
            is_option = False
            if i > 0 and lines[i-1].strip() == 'O':
                is_option = True
            item_match = re.match(r'^([A-Z]\d{6}|\d{4})$', line_content)
            if not item_match:
                i += 1
                continue
            item_code = item_match.group(0)
            if is_option:
                i += 1
                continue
            i += 1
            description_lines = []
            quantity = None
            while i < len(lines):
                line_stripped = lines[i]
                orig_line = all_lines[original_line_map[i]] if i in original_line_map else line_stripped
                if re.match(r'^([A-Z]\d{6}|\d{4})$', line_stripped):
                    i -= 1
                    break
                description_lines.append(line_stripped)
                i += 1
                if not orig_line.endswith(' '):
                    if i < len(lines):
                        next_line = lines[i]
                        qty_match = re.search(r'^(\d+)', next_line)
                        if qty_match:
                            quantity = qty_match.group(1)
                            i += 1
                    break
            if not description_lines:
                continue
            if quantity is None:
                continue
            description = " ".join(description_lines)
            if i < len(lines) and lines[i].strip().lower() in ["piece", "pièce"]:
                i += 1
            if i >= len(lines):
                continue
            unit_price_raw = lines[i].strip()
            i += 1
            unit_price = re.sub(r'[€$£¥₹¢¥₩₪₹₽]\s*', '', unit_price_raw)
            unit_price = re.sub(r'[^\d,.]', '', unit_price)
            if i < len(lines):
                i += 1
            if not re.search(r'\d', unit_price):
                continue
            item = {
                "Item": item_code,
                "Description": description,
                "Quantity": quantity,
                "UnitPrice": unit_price
            }
            items.append(item)
            item_counter += 1
        except (IndexError, AttributeError) as e:
            i += 1
    return items





def extract_table_with_pymupdf_tables(pdf_path: str) -> List[Dict[str, Any]]:
    """
    Alternative method using PyMuPDF's table detection capabilities.
    This might work better for well-structured tables.
    """
    try:
        doc = fitz.open(pdf_path)
        items = []
        item_counter = 1
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            tables = page.find_tables()
            for table in tables:
                table_data = table.extract()
                header_row = None
                data_start_idx = 0
                for i, row in enumerate(table_data):
                    if any('item' in str(cell).lower() for cell in row if cell):
                        header_row = [str(cell).lower() if cell else '' for cell in row]
                        data_start_idx = i + 1
                        break
                if header_row:
                    item_col = description_col = qty_col = price_col = -1
                    for j, header in enumerate(header_row):
                        if 'item' in header or 'product' in header:
                            item_col = j
                        elif 'description' in header:
                            description_col = j
                        elif 'quantity' in header or 'qty' in header:
                            qty_col = j
                        elif 'price' in header or 'unit' in header:
                            price_col = j
                for row in table_data[data_start_idx:]:
                    if not any(cell for cell in row):
                        continue
                    row_str = [str(cell) if cell else '' for cell in row]
                    # Only keep rows that have at least two non-empty fields (likely real items)
                    non_empty_fields = sum(1 for v in row_str if v.strip())
                    if non_empty_fields < 2:
                        continue
                    # Skip rows that look like headers or metadata
                    if any(keyword in row_str[0].lower() for keyword in ['offre', 'référence', 'client', 'identification', 'date', 'conditions', 'consultant', 'meetingroom']):
                        continue
                    if header_row and all(col >= 0 for col in [item_col, description_col, qty_col, price_col]):
                        item = {
                            "item_id": f"Item{item_counter}",
                            "Item": row_str[item_col] if item_col < len(row_str) else "",
                            "Description": row_str[description_col] if description_col < len(row_str) else "",
                            "Quantity": row_str[qty_col] if qty_col < len(row_str) else "",
                            "UnitPrice": row_str[price_col] if price_col < len(row_str) else ""
                        }
                    else:
                        item = {
                            "item_id": f"Item{item_counter}",
                            "Item": row_str[0] if len(row_str) > 0 else "",
                            "Description": row_str[1] if len(row_str) > 1 else "",
                            "Quantity": row_str[2] if len(row_str) > 2 else "",
                            "UnitPrice": row_str[3] if len(row_str) > 3 else ""
                        }
                    # Only add if at least quantity or price is present
                    if item["Quantity"].strip() or item["UnitPrice"].strip():
                        items.append(item)
                        item_counter += 1
        doc.close()
        return items
    except Exception as e:
        print(f"Warning: Table detection failed: {e}")
        return []


def main():
    parser = argparse.ArgumentParser(description='Extract table data from PDF and output as JSON')
    parser.add_argument('input_pdf', help='Path to input PDF file')
    parser.add_argument('-o', '--output', help='Output JSON file path (default: same name as input with .json extension)')
    parser.add_argument('--method', choices=['text', 'table', 'both'], default='both',
                       help='Extraction method: text parsing, table detection, or both (default: both)')
    
    args = parser.parse_args()
    
    # Validate input file
    if not Path(args.input_pdf).exists():
        print(f"Error: Input file '{args.input_pdf}' does not exist.")
        sys.exit(1)
    
    # Determine output file path
    if args.output:
        output_path = args.output
    else:
        input_path = Path(args.input_pdf)
        output_path = input_path.with_suffix('.json')
    
    print(f"Extracting data from: {args.input_pdf}")
    print(f"Output will be saved to: {output_path}")
    
    # Extract text first and save it for debugging
    text = extract_text_from_pdf(args.input_pdf)
    
    # Save raw text file for debugging
    text_output_path = Path(output_path).with_suffix('.txt')
    try:
        with open(text_output_path, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f"Raw text saved to: {text_output_path}")
    except Exception as e:
        print(f"Warning: Could not save raw text file: {e}")
    
    items = []
    used_table_method = False
    # Try table detection method first
    if args.method in ['table', 'both']:
        print("Attempting table detection...")
        items = extract_table_with_pymupdf_tables(args.input_pdf)
        used_table_method = True
    # If table detection didn't work or we want text parsing
    if (not items and args.method in ['text', 'both']) or args.method == 'text':
        if used_table_method and not items:
            print("Table detection did not find any real items, falling back to text parsing...")
        print("Using text parsing method...")
        items = parse_table_data(text)
    
    if not items:
        print("Warning: No table data found in the PDF.")
        print("This might happen if:")
        print("- The PDF doesn't contain a recognizable table structure")
        print("- The table format is different from expected")
        print("- The PDF contains scanned images instead of text")
    
    # Create output JSON structure
    output_data = {}
    for item in items:
        item_id = item.pop('item_id', f'Item{len(output_data) + 1}')
        output_data[item_id] = item
    
    # Save to JSON file
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False)
        
        print(f"Successfully extracted {len(items)} items and saved to {output_path}")
        
        # Print preview of extracted data
        if items:
            print("\nPreview of extracted data:")
            for i, (key, value) in enumerate(list(output_data.items())[:3]):
                print(f"{key}: {value}")
                if i == 2 and len(output_data) > 3:
                    print(f"... and {len(output_data) - 3} more items")
                    break
                    
    except Exception as e:
        print(f"Error saving JSON file: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()