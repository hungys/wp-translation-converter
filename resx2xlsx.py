import argparse
from ResxToXlsxConverter import ResxToXlsxConverter

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("destination", help="destination xlsx filename")
    args = parser.parse_args()

    # Load source resx files in input folder
    converter = ResxToXlsxConverter(args.destination)

    # Print all wordings and translations
    # converter.printWordings()

    # Export to xlsx file
    converter.saveXlsx()
