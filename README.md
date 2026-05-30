# Automatyzacja obsługi plików Excel — Generator raportów wynagrodzeń

Projekt zaliczeniowy demonstrujący automatyczne przetwarzanie danych z wielu plików Excel przy użyciu Pythona.

---

## Wymagania

```bash
pip install pandas openpyxl
```

Python 3.10+. Brak innych zależności zewnętrznych.

---

## Struktura projektu

```
projekt_excel/
├── data/
│   ├── dane_glowne.xlsx          # Baza pracowników
│   ├── czesc_danych_01.xlsx      # Dane miesięczne (styczeń)
│   ├── czesc_danych_02.xlsx      # Dane miesięczne (luty) — zawiera celowy błąd
│   └── czesc_danych_03.xlsx      # Dane miesięczne (marzec)
├── output/
│   └── raport_koncowy.xlsx       # Generowany raport (tworzony automatycznie)
├── tests/
│   └── test_raport.py            # Testy jednostkowe (unittest)
├── raport_generator.py           # Główny skrypt
├── create_sample_data.py         # Generator danych przykładowych
└── README.md
```

---

## Uruchomienie

```bash
# 1. Wygeneruj przykładowe dane (jeśli jeszcze nie istnieją)
python create_sample_data.py

# 2. Uruchom główny skrypt
python raport_generator.py

# 3. Uruchom testy jednostkowe
python tests/test_raport.py
```

Raport pojawi się w `output/raport_koncowy.xlsx`.

---

## Struktura plików wejściowych

### `dane_glowne.xlsx` — baza pracowników

| Kolumna           | Typ     | Opis                              |
|-------------------|---------|-----------------------------------|
| ID_pracownika     | tekst   | Unikalny identyfikator (np. P001) |
| Imie              | tekst   |                                   |
| Nazwisko          | tekst   |                                   |
| Dział             | tekst   | IT / HR / Sprzedaż / Magazyn      |
| Stanowisko        | tekst   |                                   |
| Stawka_godzinowa  | liczba  | PLN za godzinę                    |
| Prowizja_procent  | liczba  | % prowizji od sprzedaży (0 = brak)|

### `czesc_danych_NN.xlsx` — dane miesięczne

| Kolumna           | Typ     | Opis                                      |
|-------------------|---------|-------------------------------------------|
| ID_pracownika     | tekst   | Klucz łączący z plikiem bazowym           |
| Miesiac           | tekst   | Format RRRR-MM (np. 2024-01)              |
| Godziny_pracy     | liczba  | Liczba przepracowanych godzin             |
| Sprzedaz          | liczba  | Wartość sprzedaży (PLN), 0 jeśli brak     |
| Nieobecnosci_dni  | liczba  | (opcjonalnie) Liczba dni nieobecności     |

---

## Funkcjonalności

### Wczytywanie i walidacja
- Automatyczne skanowanie folderu w poszukiwaniu plików `czesc_danych_*.xlsx`
- Sprawdzanie obecności wymaganych kolumn
- Walidacja typów danych (np. czy godziny są liczbą)
- Pominięcie rekordów z ID nieistniejącym w bazie
- Zapis błędnych rekordów do arkusza „Błędy walidacji" w raporcie

### Obliczenia
- **Wynagrodzenie podstawowe** = Godziny × Stawka godzinowa
- **Prowizja** = Sprzedaż × (Prowizja_procent / 100)
- **Łączny koszt** = Wynagrodzenie podstawowe + Prowizja

### Raport wynikowy (`raport_koncowy.xlsx`)
| Arkusz                  | Zawartość                                        |
|-------------------------|--------------------------------------------------|
| Informacje              | Data, liczba pracowników, suma kosztów           |
| Dane szczegółowe        | Wszystkie rekordy miesięczne z obliczeniami      |
| Podsumowanie pracownicy | Zagregowane dane na pracownika (suma 3 miesięcy) |
| Podsumowanie działy     | Koszty według działów                            |
| Wykres                  | Wykres słupkowy kosztów wg działów               |
| Błędy walidacji         | Rekordy odrzucone podczas walidacji (jeśli są)   |

---

## Przykład użycia jako moduł

```python
from raport_generator import generuj_raport

generuj_raport(
    plik_bazowy="data/dane_glowne.xlsx",
    folder_danych="data/",
    plik_wyjsciowy="output/raport_koncowy.xlsx",
)
```

---

## Testy jednostkowe

Projekt zawiera 10 testów weryfikujących poprawność obliczeń:

```bash
python tests/test_raport.py
# Ran 10 tests in 0.1s — OK
```

Testowane scenariusze:
- Poprawność wynagrodzenia podstawowego
- Poprawność prowizji od sprzedaży
- Brak prowizji bez sprzedaży
- Poprawność łącznego kosztu
- Pominięcie rekordów z nieznanym ID
- Sumowanie godzin z wielu miesięcy
- Podział kosztów według działów
