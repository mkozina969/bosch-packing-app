# Bosch Packing List Merger – Streamlit App

Streamlit web-app za spajanje više Bosch **packing list** `.xlsx` datoteka u jedan **čisti export** s točno definiranim kolonama.

## ✨ Funkcionalnosti
- Upload 1–50 `.xlsx` datoteka odjednom (do ~200 MB po batchu).
- Automatski prepozna tablični sheet u svakoj datoteci.
- Definiraš **ciljane kolone** (comma-separated, redoslijed = redoslijed u exportu).
- Mapiranje **source → target** kolona (auto-suggest + ručno podešavanje).
- Opcija da se doda `Source_File` kolona (ime originalnog fajla).
- Skidanje gotovog **merged XLSX** fajla jednim klikom.

## 🖥 Lokalno pokretanje

1. Kloniraj repo i uđi u folder:
   ```bash
   git clone https://github.com/USERNAME/bosch-packing-app.git
   cd bosch-packing-app
