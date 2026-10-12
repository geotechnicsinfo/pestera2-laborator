# Laborator Peștera 2 — site de raportare

Site static (HTML + JavaScript, fără server) cu rezultatele de laborator geotehnic BBCGEOTECHNIC pentru amplasamentul Eoliene Peștera 2.

- `index.html` — pagina (Sinteză · Foraje · Valori caracteristice NP 122 · Date)
- `data.js` — datele, generate automat de `build_data.py`
- `build_data.py` — citește centralizatorul și fișele Excel din folderul laboratorului

## Actualizare cu probe noi
1. Pune noul centralizator (`Centralizator*.xlsx`) și fișele noi în folderul `pestera2 laborator` (fișele în `fise excel`, PDF-urile în folderul forajului).
2. În folderul `site_pestera2` rulează: `py build_data.py`
3. Încarcă noul `data.js` în repository (GitHub → Add file → Upload files).

Forfecarea directă (σ, τf, tipul încercării) se citește din buletinele PDF; pentru PDF-uri noi pe Windows instalează o dată `py -m pip install pypdf`.
