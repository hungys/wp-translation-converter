class Wording:
    def __init__(self, key):
        self.key = key
        self.translation = {}

    def addTranslation(self, value, language):
        self.translation[language] = value

    def getTranslation(self, language):
        if language in self.translation:
            return self.translation[language]
        else:
            return ""
