#!/usr/bin/python3
# -*- coding: utf8 -*-

import sys
import re
try:
    from openpyxl import load_workbook
except:
    print("python3-openpyxl csomagot telepitsd fol!")
    sys.exit(2)


class Beirasok:
    """
    Arról tárol információt, hogy bizonyos adatokat hogyan kell a NEPTUN által
    adott XLSX fájlba beírni, melyik oszlop mit jelent.

    Mindegyik adattag egy tuple, amelyiknek első eleme a NEPTUN kódot tartalmazó oszlop,
    második eleme pedig egy dict, amelyik azt mutatja, hogy az eredményt hogyan kell
    beírni. Az utóbbi dict kulcsai további oszlopok betűjelei a táblában; értékei
    pedig olyan dict-ek, amelyek azt mutatják, hogy az eredményt mutató integer
    a táblában milyen sztringgé kell változzon. Ha ehelyett None van, akkor csak sztringgé
    alakul az integer.
    """
    alairas = ("A", {
        "E": {
            0: "Megtagadva",
            1: "Aláírva"
        },
        "F": {
            0: "Megtagadva",
            1: "Aláírva"
        },
    })

    # kiírt pótdíj, ami feladatként jelenik meg a neptunban
    potdij = ("A", {
        "C": {
            0: "",
            1: "1(egy)",
            2: "2(kettő)",
        },
    })

    # kzh, nzh, nhf, ... pontszám
    pontszam = ("A", {
        "C": None,
    })

    imscpont = ("A", {
        "H": None,
    })

    gonogo = ("A", {
        "C": {
            0: "Nem felelt meg",
            1: "Megfelelt"
        }
    })

    felevkozi = ("A", {
        "E": {
            0: "Nem teljesítette",
            1: "Elégtelen",
            2: "Elégséges",
            3: "Közepes",
            4: "Jó",
            5: "Jeles"
        },
        "F": {
            0: "Nem teljesítette",
            1: "Elégtelen",
            2: "Elégséges",
            3: "Közepes",
            4: "Jó",
            5: "Jeles"
        },
    })

    vizsga = ("C", {
        "H": {
            0: "",
            1: "Elégtelen",
            2: "Elégséges",
            3: "Közepes",
            4: "Jó",
            5: "Jeles"
        },
        "I": {
            0: "",
            1: "Elégtelen",
            2: "Elégséges",
            3: "Közepes",
            4: "Jó",
            5: "Jeles"
        },
        "J": {
            0: "Igen",
            1: "Nem",
            2: "Nem",
            3: "Nem",
            4: "Nem",
            5: "Nem"
        },
    })


def csv_beolvas(fajlnev):
    """
    Beolvassa a CSV fájlt, amelyben minden sorban egy NEPTUN és egy eredmény van.
    Paraméter:
        fajlnev (str): fájlnév
    Vissza:
        (dict): neptun -> eredmény
    """
    file = open(fajlnev, "r")
    eredmenyek = {}
    for line in file:
        split = line.strip().replace(",", "\t").replace(";", "\t").split("\t")
        neptun = split[0]
        eredmeny = split[1]
        if eredmeny == "":
            eredmeny = "0"
        eredmenyek[neptun] = int(eredmeny)
    return eredmenyek


def jegyeket_beir(worksheet, eredmenyek, beiras):
    """
    Beírja egy megnyitott worksheetbe a CSV-ből beolvasott adatokat.
    Paraméterek:
        worksheet (Worksheet)
        eredmenyek (dict): neptun -> eredmény.
        beiras (tuple): lásd a Beiras osztályt.
    """
    def oszlopindex(oszlop):
        return ord(oszlop) - ord("A")

    # néha csak neptun kód van, néha meg "név ( neptun )"
    neptunextractor = re.compile(r".*?\(\s*([A-Z0-9]{6})\s*\)")
    def neptunextract(s):
        s = str(s)
        match = neptunextractor.match(s)
        if match == None:
            return s
        else:
            return match.group(1)

    (neptunoszlop, eredmenyoszlopok) = beiras
    neptunoszlopidx = oszlopindex(neptunoszlop)
    for row in worksheet.rows:
        neptun = neptunextract(row[neptunoszlopidx].value)
        if neptun not in eredmenyek:
            # lehet nincs ilyen neptun, mert nem ez a kurzus, vagy ez fejléc oszlopot találtuk meg
            continue
        eredmeny = eredmenyek[neptun]
        for oszlopnev, dikt in eredmenyoszlopok.items():
            if dikt is not None:
                eredmenystr = dikt[eredmeny]
            else:
                eredmenystr = str(eredmeny)
            eredmenyoszlopidx = oszlopindex(oszlopnev)
            row[eredmenyoszlopidx].value = eredmenystr


def xlsx_feldolgoz(inputfajlnev, csvfajlnev, outputfajlnev, beiras):
    """
    XLSX + CSV beolvasása, XLSX új fájlba írása.
    """
    eredmenyek = csv_beolvas(csvfajlnev)
    workbook = load_workbook(inputfajlnev)
    worksheet = workbook.active
    jegyeket_beir(worksheet, eredmenyek, beiras)
    workbook.save(outputfajlnev)


class Args:
    """
    Parancssori argumentumokat parsol be.
    """
    inputfajlnev = "neptun_orig.xlsx"
    outputfajlnev = "neptun.xlsx"
    csvfajlnev = "marks.txt"
    beirastipus = Beirasok.vizsga
    debug = False

    def __init__(self, argv):
        i = 1
        while i < len(argv):
            arg = argv[i]
            if arg == "-f":
                i += 1
                self.inputfajlnev = argv[i]
            elif arg == "-o":
                i += 1
                self.outputfajlnev = argv[i]
            elif arg == "-m":
                i += 1
                self.csvfajlnev = argv[i]
            elif arg == "-d":
                self.debug = True
            elif arg == "-tv":
                self.beirastipus = Beirasok.vizsga
            elif arg == "-tf":
                self.beirastipus = Beirasok.felevkozi
            elif arg == "-ta":
                self.beirastipus = Beirasok.alairas
            elif arg == "-tk":
                self.beirastipus = Beirasok.pontszam
            elif arg == "-ti":
                self.beirastipus = Beirasok.imscpont
            elif arg == "-tn":
                self.beirastipus = Beirasok.gonogo
            elif arg == "-tp":
                self.beirastipus = Beirasok.potdij
            elif arg == "-h":
                print("%s -f neptun_orig.xlsx -o neptun.xlsx -m marks.txt [-ta|-tf|-tv|-tk|-ti|-tn|-tp]\nAz input fajl kiterjesztese legyen .xlsx!" % (argv[0]))
                sys.exit(0)
            else:
                print("Ervenytelen argumentum: %s" % (arg))
            i += 1


try:
    args = Args(sys.argv)
    xlsx_feldolgoz(args.inputfajlnev, args.csvfajlnev, args.outputfajlnev, args.beirastipus)
except Exception as e:
    print(e)
    sys.exit(1)
