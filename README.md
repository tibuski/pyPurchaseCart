# pyPurchaseCart

A Python tool to extract item tables from PDF sales quotes and output them as JSON files.

## Features
- Extracts tables with item codes, descriptions, quantities, and unit prices from PDFs.
- Supports both table detection and text parsing methods.
- Outputs structured JSON for easy integration with other systems.

## Requirements
- Python 3.7+
- [PyMuPDF](https://pymupdf.readthedocs.io/) (`pip install PyMuPDF`)

## Usage

```sh
python main.py <input_pdf> [-o <output_json>] [--method text|table|both]
```

- `<input_pdf>`: Path to the PDF file to process.
- `-o <output_json>`: (Optional) Output JSON file path. Defaults to the same name as the PDF with `.json` extension.
- `--method`: Extraction method. `text` for text parsing, `table` for table detection, `both` (default) tries table detection first, then text parsing if needed.

### Example

```sh
python main.py data/PS2500001\ MEETINGROOM\ MERKSEM\ 1.13.pdf
```

## Output
The script will create a JSON file with extracted items, e.g.:

```json
{
  "Item1": {
    "Item": "A103970",
    "Description": "SAMSUNG QM85C 85-inch 4K/UHD Digital Signage LED Display ...",
    "Quantity": "1",
    "UnitPrice": "1975,00"
  },
  ...
}
```

## License
MIT
