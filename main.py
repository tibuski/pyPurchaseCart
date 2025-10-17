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
    Looks for items with A+6digit codes and extracts Item, Description, Quantity, UnitPrice.
    Handles multi-line descriptions using trailing space detection.
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
    
    # Find table boundaries: starts with first A+6digits, ends at clear end markers
    table_start_idx = None
    table_end_idx = None
    
    for i, line in enumerate(lines):
        line_stripped = line.strip()
        
        # Look for first line that is exactly A+6digits pattern
        if table_start_idx is None and re.match(r'^[A-Z]\d{6}$', line_stripped):
            table_start_idx = i
            print(f"Found table start at line {i}: {line_stripped}")
            continue
        
        # After we found the start, look for first line that clearly indicates end of items
        # (but skip empty lines and description/quantity/price lines that are part of items)
        if table_start_idx is not None and line_stripped:
            # Check if this could be an item code (A+6digits exactly)
            if re.match(r'^[A-Z]\d{6}$', line_stripped):
                continue  # This is another item, keep going
            
            # Check if this could be part of an item (description, quantity, unit, price)
            # Skip lines that are likely part of item data
            if (re.match(r'^[1-9]\d{0,2}$', line_stripped) or  # quantity
                line_stripped.lower() == 'piece' or  # unit
                re.search(r'\d+[,.]?\d*', line_stripped) or  # contains price-like numbers
                len(line_stripped) > 10):  # long text (likely description)
                continue
            
            # Look for clear end markers
            if any(keyword in line_stripped.lower() for keyword in 
                   ['subtotal', 'total', 'delivery', 'vat', 'payment', 'terms']):
                table_end_idx = i
                print(f"Found table end at line {i}: {line_stripped}")
                break
    
    if table_start_idx is None:
        print("Warning: Could not find any item codes with A+6digits pattern")
        return items
    
    if table_end_idx is None:
        print("Warning: Could not find table end, processing to end of document")
        table_end_idx = len(lines)
    
    print(f"Processing table data from line {table_start_idx} to {table_end_idx - 1}")
    
    # Parse items starting from the first item code
    i = table_start_idx
    item_counter = 1
    
    # Parse each item by looking for item codes and collecting data until next item
    while i < table_end_idx and i < len(lines):
        try:
            # Look for item code (A+6digits pattern)
            line_content = lines[i].strip()
            
            # Check if previous line was "O" (indicating this is an option item to skip)
            is_option = False
            if i > 0 and lines[i-1].strip() == 'O':
                is_option = True
            
            # Find A+6digits pattern
            item_match = re.match(r'^[A-Z]\d{6}$', line_content)
            if not item_match:
                i += 1
                continue
                
            item_code = item_match.group(0)
            
            # Skip option items
            if is_option:
                i += 1
                continue
            i += 1
            
            # Collect description lines using the trailing space rule
            description_lines = []
            quantity = None
            
            while i < table_end_idx and i < len(lines):
                line_stripped = lines[i]  # Already stripped
                # Get original line for space checking
                orig_line = all_lines[original_line_map[i]] if i in original_line_map else line_stripped
                
                # If we find another item code, we went too far
                if re.match(r'^[A-Z]\d{6}$', line_stripped):
                    i -= 1  # Back up one line
                    break
                
                # Add this line to description
                description_lines.append(line_stripped)
                i += 1
                
                # If original line doesn't end with space, this is the last line of description
                # The next line should be quantity
                if not orig_line.endswith(' '):
                    # Check if next line is quantity (handle various quantity formats)
                    if i < table_end_idx and i < len(lines):
                        next_line = lines[i]  # Already stripped
                        # Extract number from quantity line (handle "1", "4 pièce", "2 pieces", etc.)
                        qty_match = re.search(r'^(\d+)', next_line)
                        if qty_match:
                            quantity = qty_match.group(1)
                            i += 1  # Move past quantity
                    break
            
            if not description_lines:
                continue
                
            if quantity is None:
                continue
                
            description = " ".join(description_lines)
            
            # Skip unit (should be "piece" or "pièce")
            if i < len(lines) and lines[i].strip().lower() in ["piece", "pièce"]:
                i += 1
            
            # Get unit price
            if i >= len(lines):
                continue
                
            unit_price_raw = lines[i].strip()
            i += 1
            
            # Clean unit price: remove currency symbols, spaces, keep only numbers and decimal separators
            unit_price = re.sub(r'[€$£¥₹¢¥₩₪₹₽]\s*', '', unit_price_raw)  # Remove common currency symbols
            unit_price = re.sub(r'[^\d,.]', '', unit_price)  # Keep only digits, commas, and dots
            
            # Skip total price (next line)
            if i < len(lines):
                i += 1
            
            # Validate unit price (should contain numbers)
            if not re.search(r'\d', unit_price):
                continue
            
            item = {
                "Item": item_code,
                "Description": description,
                "Quantity": quantity,
                "UnitPrice": unit_price  # Keep original format with comma for decimals
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
            
            # Try to find tables on the page
            tables = page.find_tables()
            
            for table in tables:
                table_data = table.extract()
                
                # Look for header row
                header_row = None
                data_start_idx = 0
                
                for i, row in enumerate(table_data):
                    if any('item' in str(cell).lower() for cell in row if cell):
                        header_row = [str(cell).lower() if cell else '' for cell in row]
                        data_start_idx = i + 1
                        break
                
                # If we found a header, map columns
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
                
                # Extract data rows
                for row in table_data[data_start_idx:]:
                    if not any(cell for cell in row):  # Skip empty rows
                        continue
                    
                    row_str = [str(cell) if cell else '' for cell in row]
                    
                    # Skip total/subtotal rows
                    if any('total' in cell.lower() for cell in row_str):
                        continue
                    
                    if header_row and all(col >= 0 for col in [item_col, description_col, qty_col, price_col]):
                        # Use column mapping
                        item = {
                            "item_id": f"Item{item_counter}",
                            "Item": row_str[item_col] if item_col < len(row_str) else "",
                            "Description": row_str[description_col] if description_col < len(row_str) else "",
                            "Quantity": row_str[qty_col] if qty_col < len(row_str) else "",
                            "UnitPrice": row_str[price_col] if price_col < len(row_str) else ""
                        }
                    else:
                        # Fallback to positional parsing
                        item = {
                            "item_id": f"Item{item_counter}",
                            "Item": row_str[0] if len(row_str) > 0 else "",
                            "Description": row_str[1] if len(row_str) > 1 else "",
                            "Quantity": row_str[2] if len(row_str) > 2 else "",
                            "UnitPrice": row_str[3] if len(row_str) > 3 else ""
                        }
                    
                    if any(item[key] for key in ["Item", "Description", "Quantity", "UnitPrice"]):
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
    
    items = []
    
    # Try table detection method first
    if args.method in ['table', 'both']:
        print("Attempting table detection...")
        items = extract_table_with_pymupdf_tables(args.input_pdf)
    
    # If table detection didn't work or we want text parsing
    if (not items and args.method in ['text', 'both']) or args.method == 'text':
        print("Using text parsing method...")
        text = extract_text_from_pdf(args.input_pdf)
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