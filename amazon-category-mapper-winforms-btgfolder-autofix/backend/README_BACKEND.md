# Backend

Pokretanje ručno:

```bat
cd backend
run_backend.bat
```

Health check:

```text
http://127.0.0.1:8008/health
```

## Endpoints

### POST `/api/import-category-catalog`

Uvoz jedne BTG/category catalog datoteke u lokalnu SQLite bazu.

Multipart fields:

```text
category_file = .xls/.xlsx/.csv BTG ili catalog file
marketplace = AUTO/FR/IT/ES/DE/... opcionalno
```

### POST `/api/import-category-folder`

Uvoz cijelog lokalnog foldera s BTG datotekama.

JSON body:

```json
{
  "folder_path": "C:\\AmazonBTG",
  "marketplace": "AUTO"
}
```

Ako je marketplace `AUTO`, backend pokušava zaključiti marketplace iz sheet namea, file namea ili parent foldera (`FR`, `IT`, `ES`...).

### POST `/api/map-csv`

Mapira ulazni CSV/XLSX pomoću:

```text
input_file
mapping_file
category_file opcionalno
marketplaces = FR,IT,ES
```

Ako si BTG već uvezao u bazu, `category_file` više ne moraš slati svaki put.
