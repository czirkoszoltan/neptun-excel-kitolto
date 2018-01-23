# NEPTUN Excel kitöltő

NEPTUN-ból származó XLSX fájlokba ír be CSV-ből adatokat.

Használat:
```
neptun_beiro.py -f input.xlsx -o output.xlsx -m marks.txt [-ta|-tf|-tv|-tk|-ti|-tn|-tp]
```

Paraméterek:

- `-f input.xlsx` – az üres XLSX fájl. Kötelező az `.xlsx` kiterjesztés.

- `-o output.xlsx` – a kitöltött XLSX fájl.

- `-m marks.txt` – a beírandó adatok, `NEPTUNKÓD EREDMÉNY` formátumban soronként.

- `-ta`: aláírás, az eredmények értéke 0 és 1 lehet.

- `-tf`: félévközi jegy beírása, 0 = nem teljesítette, 1...5 = jegy.

- `-ti`: IMSc pont beírása; ez a félévközi jegy beírása után használható.

- `-tv`: vizsgajegy beírása, 0 = nem jelent meg, 1...5 = jegy.

- `-tk`: pontszám beírása.

- `-tn`: megfelelt vagy nem megfelelt eredmény beírása.

- `-tp`: pótdíj beírása; az eredmény a kiírt pótdíjak száma.
