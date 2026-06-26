# Amazon Category Mapper - WinForms + Python backend + BTG folder importer

Ova verzija je prilagođena workflowu gdje imaš:

- ulazni CSV/XLSX koji može imati samo `NODE ID`, ili puni format s `product type`, `Amazon category name`, `EAN` itd.
- Amazon European Browse Node Mapping Table s tabom `MAPPINGS`
- puno Amazon BTG datoteka po marketplaceu i kategoriji, npr. grocery, garden, tools, home...

## Što radi

Za FR/IT/ES radi sigurno mapiranje bez izmišljanja:

```text
Ulazni DE/source NODE ID
  ↓
European Browse Node Mapping Table / MAPPINGS
  ↓
Node Id in FR / Node Id in IT / Node Id in ES
  ↓
lokalna SQLite baza s uvezenim BTG kategorijama
  ↓
FR / IT / ES stvarni Amazon category path
```

Ako u BTG bazi nema target nodea, program NE izmišlja naziv, nego stavlja status:

```text
DIRECT_NODE_MAPPING_NAME_MISSING
```

## Je li dovoljno da input ima samo NODE ID?

Da, za FR/IT/ES je dovoljno da ulazni CSV/XLSX ima barem stupac:

```text
NODE ID
```

Primjer minimalnog CSV-a:

```csv
NODE ID
3597051031
316968011
```

Program će iz `MAPPINGS` taba dohvatiti `Node Id in FR`, `Node Id in IT`, `Node Id in ES`, a zatim iz uvezene BTG baze dohvatiti nazive/pathove kategorija.

## Kako organizirati BTG datoteke

Preporuka:

```text
C:\AmazonBTG
  FR
    fr_grocery.xls
    fr_garden.xls
    fr_tools.xls
  IT
    it_grocery.xls
    it_garden.xls
    it_tools.xls
  ES
    es_grocery.xls
    es_garden.xls
    es_tools.xls
  DE
    de_food.xls
    de_garden.xls
```

Možeš imati puno kategorija. Alat rekurzivno čita cijeli folder i sve podfoldere.

BTG file obično mora imati stupce:

```text
Node ID | Node Path | Refinement Link
```

## Pokretanje

1. Raspakiraj ZIP.
2. Otvori `winforms/AmazonCategoryMapperWinForms/AmazonCategoryMapperWinForms.csproj` u Visual Studio 2022.
3. Pokreni WinForms aplikaciju.
4. Klikni `Pokreni Python backend`.
5. U `BTG folder` odaberi folder gdje su svi BTG fileovi.
6. Marketplace ostavi `AUTO` ako su fileovi/sheetovi nazvani npr. `fr-grocery`, `it-garden`, `es-tools`, ili su u folderima `FR`, `IT`, `ES`.
7. Klikni `Uvezi cijeli BTG folder`.
8. Odaberi ulazni CSV/XLSX.
9. Odaberi Amazon European Browse Node Mapping Excel (`MAPPINGS` tab).
10. Označi FR/IT/ES.
11. Klikni `Mapiraj CSV`.
12. Spremi novi CSV.

## Za NL/PL/IE/SE

Za te države trenutni European mapping file nema direktne stupce, pa ih alat ne može službeno mapirati iz tog Excela.

Ako uključiš OpenAI fallback, alat može pokušati dati preporuku, ali status ostaje za pregled (`AI_NEEDS_REVIEW` / `NEED_REVIEW`). To nije službeni potvrđeni rezultat dok ga korisnik ne potvrdi.

## Lokalna baza

Sve uvezene BTG kategorije spremaju se u:

```text
backend/data/learning.db
```

Tablica:

```text
category_catalog
```

Ključni podaci:

```text
marketplace | node_id | node_path | category_name | source_file
```

Kasnije se ova logika može prebaciti na MySQL.

## Nema pandas/numpy

Backend namjerno ne koristi pandas/numpy zbog Windows Application Control policy problema. Koristi samo:

```text
openpyxl
xlrd
sqlite3
standardni Python
```


## Ažuriranje: automatski BTG folder import prije mapiranja

Ako je u WinForms aplikaciji odabran `BTG folder`, backend ga sada automatski uveze/azurira prije mapiranja.
To znači da je dovoljno:

1. Odabrati folder s BTG datotekama, npr. `C:\AmazonBTG`.
2. Odabrati ulazni CSV/XLSX.
3. Odabrati European Browse Node Mapping Table.
4. Kliknuti `Mapiraj CSV`.

Možeš i dalje ručno kliknuti `Uvezi cijeli BTG folder`, ali više nije obavezno ako je folder upisan/odabran.

Primjer: za source `NODE ID = 4288535031`, European mapping daje `Node Id in FR = 4338712031`.
Ako je učitan FR garden BTG, alat popunjava:

`FR = Jardin > Jardinage > Protection et anti-nuisibles pour jardin > Lutte contre les mauvaises herbes > Toiles de paillage`

Ako želiš `Bordures pour jardin`, to je drugi FR node (`4338588031`), pa to nije isti direktni European mapping rezultat. Takav slučaj treba ručno ispraviti ili kasnije raditi kao AI/smart suggestion.
