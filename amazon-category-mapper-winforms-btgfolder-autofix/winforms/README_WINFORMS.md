# WinForms aplikacija

Otvori `AmazonCategoryMapperWinForms.csproj` u Visual Studio 2022.

Redoslijed:

1. Pokreni Python backend.
2. Odaberi `BTG folder` i klikni `Uvezi cijeli BTG folder`.
3. Odaberi ulazni CSV/XLSX. Za osnovno mapiranje dovoljan je stupac `NODE ID`.
4. Odaberi Amazon European Browse Node Mapping Excel s `MAPPINGS` tabom.
5. Označi `FR`, `IT`, `ES`.
6. Klikni `Mapiraj CSV`.
7. Ručno ispravi ako treba.
8. Za ispravljene vrijednosti postavi `FR_status`, `IT_status`, itd. na `CORRECTED` ili `CONFIRMED`.
9. Klikni `Spremi ispravke u learning bazu`.
10. Spremi novi CSV.
