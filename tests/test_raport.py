"""
Testy jednostkowe dla raport_generator.py
Uruchomienie: python tests/test_raport.py
lub po zainstalowaniu pytest: pytest tests/test_raport.py -v
"""

import sys
import os
import unittest
import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from raport_generator import scal_i_oblicz, podsumowanie_dzialowe


def make_baza():
    return pd.DataFrame({
        "ID_pracownika": ["P001", "P002", "P003"],
        "Imie": ["Anna", "Bartosz", "Celina"],
        "Nazwisko": ["Kowalska", "Nowak", "Wiśniewska"],
        "Dział": ["IT", "IT", "Sprzedaż"],
        "Stanowisko": ["Programista", "Analityk", "Handlowiec"],
        "Stawka_godzinowa": [100.0, 80.0, 60.0],
        "Prowizja_procent": [0.0, 0.0, 5.0],
    })


def make_dane():
    return pd.DataFrame({
        "ID_pracownika": ["P001", "P002", "P003"],
        "Miesiac": ["2024-01", "2024-01", "2024-01"],
        "Godziny_pracy": [160.0, 160.0, 160.0],
        "Sprzedaz": [0.0, 0.0, 10000.0],
        "Nieobecnosci_dni": [0, 0, 0],
    })


class TestObliczenia(unittest.TestCase):

    def test_wynagrodzenie_podstawowe(self):
        sz, _ = scal_i_oblicz(make_baza(), make_dane())
        row = sz[sz["ID_pracownika"] == "P001"].iloc[0]
        self.assertAlmostEqual(row["Wynagrodzenie_podstawowe"], 16000.0)

    def test_prowizja(self):
        sz, _ = scal_i_oblicz(make_baza(), make_dane())
        row = sz[sz["ID_pracownika"] == "P003"].iloc[0]
        self.assertAlmostEqual(row["Prowizja"], 500.0)  # 5% z 10 000

    def test_brak_prowizji_bez_sprzedazy(self):
        sz, _ = scal_i_oblicz(make_baza(), make_dane())
        row = sz[sz["ID_pracownika"] == "P001"].iloc[0]
        self.assertAlmostEqual(row["Prowizja"], 0.0)

    def test_laczny_koszt(self):
        sz, _ = scal_i_oblicz(make_baza(), make_dane())
        row = sz[sz["ID_pracownika"] == "P003"].iloc[0]
        self.assertAlmostEqual(row["Laczny_koszt"], 160 * 60 + 500)

    def test_nieznane_id_pomijane(self):
        dane = pd.concat([
            make_dane(),
            pd.DataFrame([{"ID_pracownika": "X999", "Miesiac": "2024-01",
                           "Godziny_pracy": 100.0, "Sprzedaz": 0.0, "Nieobecnosci_dni": 0}])
        ], ignore_index=True)
        sz, _ = scal_i_oblicz(make_baza(), dane)
        self.assertNotIn("X999", sz["ID_pracownika"].values)

    def test_suma_godzin_w_podsumowaniu(self):
        sz, podsum = scal_i_oblicz(make_baza(), make_dane())
        self.assertAlmostEqual(podsum["Godziny_pracy"].sum(), sz["Godziny_pracy"].sum())

    def test_dwa_miesiace_sumowane(self):
        dane = pd.DataFrame({
            "ID_pracownika": ["P001", "P001"],
            "Miesiac": ["2024-01", "2024-02"],
            "Godziny_pracy": [160.0, 168.0],
            "Sprzedaz": [0.0, 0.0],
        })
        _, podsum = scal_i_oblicz(make_baza(), dane)
        row = podsum[podsum["ID_pracownika"] == "P001"].iloc[0]
        self.assertAlmostEqual(row["Godziny_pracy"], 328.0)


class TestPodsumowanieDzialowe(unittest.TestCase):

    def test_liczba_dzialow(self):
        sz, _ = scal_i_oblicz(make_baza(), make_dane())
        dzial = podsumowanie_dzialowe(sz)
        self.assertEqual(len(dzial), 2)

    def test_suma_kosztow_it(self):
        sz, _ = scal_i_oblicz(make_baza(), make_dane())
        dzial = podsumowanie_dzialowe(sz)
        it = dzial[dzial["Dział"] == "IT"].iloc[0]
        self.assertAlmostEqual(it["Laczny_koszt"], 160 * 100 + 160 * 80)

    def test_suma_dzialow_equals_total(self):
        sz, _ = scal_i_oblicz(make_baza(), make_dane())
        dzial = podsumowanie_dzialowe(sz)
        self.assertAlmostEqual(dzial["Laczny_koszt"].sum(), sz["Laczny_koszt"].sum())


if __name__ == "__main__":
    unittest.main(verbosity=2)
