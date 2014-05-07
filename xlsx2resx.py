import argparse
import os
import sys
from XlsxToResxConverter import XlsxToResxConverter

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("source", help="source xlsx filename")
    args = parser.parse_args()

    if os.path.isfile(args.source) == False:
        print("Error: source file not found")
        sys.exit()

    # Load source xlsx file
    converter = XlsxToResxConverter(args.source)

    # Print all wordings and translations
    # converter.printWordings()

    # Export to resx files
    converter.saveResx()
