# csv2xlsx
converter from CSV to XLSX

## Usage

```bash
$ python3 csv2xlsx.py --help
usage: csv2xlsx.py [-h] [-o OUTPUT] [-e ENCODING] [-s SHEET_NAME] [--csv-field-size-limit CSV_FIELD_SIZE_LIMIT] input

positional arguments:
  input                 input CSV file name

optional arguments:
  -h, --help            show this help message and exit
  -o OUTPUT, --output OUTPUT
                        output XLSX file name
  -e ENCODING, --encoding ENCODING
                        encoding type of input file
  -s SHEET_NAME, --sheet-name SHEET_NAME
                        sheet name in output file
  --csv-field-size-limit CSV_FIELD_SIZE_LIMIT
                        limit of CSV file size
```
