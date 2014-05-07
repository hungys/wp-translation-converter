wp-translation-converter
========================

A tiny script to convert between resx and xlsx for Windows Phone localization resource file

# Requirement

* Python 3.3+
* xlrd 0.9.3: use `pip3 install xlrd` to install
* XlsxWriter 0.5.5: use `pip3 install XlsxWriter` to install

# Resx to Xlsx

convert resx resource file to Excel worksheet

- source resx files (e.g. AppResources.resx, AppResources.zh-TW.resx...) should be placed in **input** folder
- usage: `python3 resx2xlsx.py [destination_file]`
- core function: ResxToXlsxConverter.py

# Xlsx to Resx

convert Excel worksheet to resx resource file

- usage: `python3 xlsx2resx.py [source_file]`
- output resx files (e.g. AppResources.resx, AppResources.zh-TW.resx...) will be saved in **output** folder
- core function: XlsxToResxConverter.py
