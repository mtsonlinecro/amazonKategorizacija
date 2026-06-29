# WinForms aplikacija

UI je pojednostavljen:

1. Pokreni Python backend.
2. Odaberi samo ulazni CSV/XLSX/TXT.
3. Odaberi države.
4. Klikni Mapiraj.

Više nema ručnog odabira:

- Amazon MAPPINGS Excela
- jednog category catalog/BTG filea
- BTG foldera
- OpenAI API keya

Backend sam traži `BTG` folder i `mapping_file.xls/xlsx` u rootu projekta. Ako ih ne pronađe, status/greška će to javiti.


## Novo u UI-ju

Dodani su:

- progress bar za trenutno mapiranje
- tekstualni status obrade
- gumb `Obriši BTG cache`

`Obriši BTG cache` briše samo uvezene BTG/category podatke. Ne briše tvoje ručne ispravke koje si spremio preko `Spremi ručne ispravke`.
