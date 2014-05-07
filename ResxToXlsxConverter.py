from Wording import Wording
from xlsxwriter.workbook import Workbook
import xml.etree.ElementTree as ET
import os

class ResxToXlsxConverter:
    def __init__(self, path="wording.xlsx"):
        self.path = path
        self.document = Workbook(path)
        self.sheet = self.document.add_worksheet()
        self.languages = []
        self.wordings = {}

        self.loadResx()

    def addWording(self, key, value, language):
        if language not in self.languages:
            self.languages.append(language)

        if key in self.wordings:
            self.wordings[key].addTranslation(value, language)
        else:
            self.wordings[key] = Wording(key)
            self.wordings[key].addTranslation(value, language)

    def printWordings(self):
        for key, value in self.wordings.items():
            print("{0}: {1}\n".format(key, value.translation))

    def loadResx(self):
        if not os.path.exists("input"):
            os.makedirs("input")

        filelist = [f for f in os.listdir("input")]
        for f in filelist:
            if f == "AppResources.resx":
                language = "en-US"
            else:
                language = f.split(".")[1]

            resx_tree = ET.parse(os.path.join("input", f)).findall("data")

            print(len(resx_tree), "wordings found in", os.path.join("input", f))

            for unit in resx_tree:
                self.addWording(unit.attrib["name"], unit.find("value").text, language)

    def saveXlsx(self):
        cur_row = 1
        self.sheet.write(0, 0, "Resource ID")

        if "en-US" in self.languages:
            self.languages.insert(0, self.languages.pop(self.languages.index("en-US")))
        if "zh-TW" in self.languages:
            self.languages.insert(1, self.languages.pop(self.languages.index("zh-TW")))

        for language in self.languages:
            self.sheet.write(0, self.languages.index(language) + 1, language)

        for key, wording in sorted(self.wordings.items()):
            self.sheet.write(cur_row, 0, wording.key)
            for language in self.languages:
                self.sheet.write(cur_row, self.languages.index(language) + 1, wording.getTranslation(language))
            cur_row = cur_row + 1

        format = self.document.add_format()
        format.set_pattern(1)
        format.set_bg_color("orange")
        format.set_bold()
        self.sheet.set_column(0, 0, 30)
        self.sheet.set_column(1, len(language), 60)
        self.sheet.set_row(0, cell_format=format)

        self.document.close()

        print("Languages exported: ", self.languages)
        print(self.path, "saved")
