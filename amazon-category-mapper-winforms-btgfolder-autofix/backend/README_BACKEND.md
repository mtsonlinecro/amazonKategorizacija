# Backend

Aktivni backend je samo:

```text
server.py
```

Stara FastAPI/mapper verzija je maknuta da ne zbunjuje. Ovaj backend ne koristi pandas/numpy i radi s običnim `http.server` API-jem na:

```text
http://127.0.0.1:8008
```

## Automatske putanje

Backend automatski traži:

- `BTG` folder u rootu projekta
- `mapping_file.xls/xlsx` ili `*mapping*.xls*` u rootu projekta

Opcionalno override:

```text
BTG_FOLDER_PATH=C:\putanja\do\BTG
MAPPING_FILE_PATH=C:\putanja\do\mapping_file.xls
```

## AI

AI je ugašen po defaultu. Za opcionalni fallback izmijeni:

```text
ai_settings.py
```


## Progress i cache endpointi

WinForms za svako mapiranje šalje `job_id`, a backend vraća stanje na:

```text
GET /api/progress?job_id=...
```

BTG/category cache se može obrisati bez brisanja ručnih learning ispravaka:

```text
POST /api/clear-cache
```

Cache se i automatski osvježava ako se BTG datoteka promijeni po veličini ili modified datumu.

## Opcionalni translate provider za NL/PL/IE/SE

Za države bez direktnog Amazon European Browse Node Mappinga program sada može prevoditi samo zadnji dio njemačke kategorije, npr. `Rasenkanten`, u ciljni jezik i zatim tražiti rezultat u istoj BTG obitelji.

Ključevi se podešavaju u:

```txt
backend/api_keys.py
```

Najjednostavnije je postaviti DeepL:

```python
TRANSLATION_ENABLED = True
TRANSLATION_PROVIDER = "deepl"
DEEPL_API_KEY = "OVDJE_API_KEY"
DEEPL_USE_FREE_API = True
```

Ako ne želiš ručno birati provider, ostavi:

```python
TRANSLATION_PROVIDER = "auto"
```

Tada se koristi prvi provider za koji je postavljen API ključ. Podržani provider-i su: DeepL, Google Cloud Translation, Azure Translator i OpenAI.

Prijevodni cache je u:

```txt
backend/cache/translation_cache.json
```

Cache sprema samo kratke prijevode leaf pojmova, ne BTG datoteke. Može se obrisati iz WinForms aplikacije gumbom **Obriši translate cache**.
