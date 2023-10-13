class Aktie():
    def __init__(self, name, wkn, isin, branche, sektor):
        self.name = name
        self.wkn = wkn
        self.isin = isin
        self.branche = branche
        self.sektor = sektor
        self.daten = []

    def setdata(self, data):
        self.daten = data