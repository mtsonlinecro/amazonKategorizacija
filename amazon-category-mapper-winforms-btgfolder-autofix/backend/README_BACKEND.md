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
