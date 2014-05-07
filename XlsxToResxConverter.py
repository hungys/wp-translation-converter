from Wording import Wording
import xlrd
import os

class XlsxToResxConverter:
    def __init__(self, path):
        self.document = xlrd.open_workbook(path)
        self.sheet = self.document.sheet_by_index(0)
        self.languages = self.sheet.row_values(0)[1:]
        self.wordings = {}

        print("Source file:", path)
        print(len(self.languages), "languages loaded:", self.languages)
        print(self.sheet.nrows - 1, "wordings found")

        self.loadXlsx()

    def addWording(self, key, value, language):
        if key in self.wordings:
            self.wordings[key].addTranslation(value, language)
        else:
            self.wordings[key] = Wording(key)
            self.wordings[key].addTranslation(value, language)

    def printWordings(self):
        for key, value in self.wordings.items():
            print("{0}: {1}\n".format(key, value.translation))

    def loadXlsx(self):
        for i in range(1, self.sheet.nrows):
            row = self.sheet.row_values(i)
            for language in self.languages:
                self.addWording(row[0], row[self.languages.index(language) + 1], language)

    def saveResx(self):
        if not os.path.exists("output"):
            os.makedirs("output")

        filelist = [f for f in os.listdir("output")]
        for f in filelist:
            os.remove(os.path.join("output", f))

        for language in self.languages:
            filename = "AppResources.resx" if language=="en-US" else "AppResources.{0}.resx".format(language)
            resx = open(os.path.join("output", filename), "w+", encoding="UTF-8")
            template = open("ResxTemplate.resx", "r", encoding="UTF-8").read()
            output = ""
            for key, value in self.wordings.items():
                output += "  <data name=\"{0}\" xml:space=\"preserve\">\r\n".format(key)
                output += "    <value>{0}</value>\r\n".format(value.translation[language])
                output += "  </data>\r\n"


            resx.write(template.replace("<!--- output -->", output))
            resx.close()

            print(os.path.join("output", filename), "generated")
