"""
Konfiguracja automatyzacji Płatnika
Dostosuj wartości do swojego środowiska
"""
from pathlib import Path

# === ŚCIEŻKI ===
BASE_DIR = Path(__file__).parent
EXCEL_INPUT = BASE_DIR / "dane_pracownicy.xlsx"
KEDU_OUTPUT_DIR = BASE_DIR / "kedu"
LOG_DIR = BASE_DIR / "logs"
CONFIRMATION_DIR = BASE_DIR / "output"

# === PŁATNIK ===
PLATNIK_EXE = r"C:\Program Files (x86)\Platnik\platnik.exe"
PLATNIK_WINDOW_TITLE = "Płatnik"
PLATNIK_TIMEOUT = 30  # sekundy oczekiwania na reakcję GUI

# === DANE PŁATNIKA (ZUS) ===
# !! UZUPEŁNIJ PONIŻSZE DANE — BEZ NICH KEDU NIE ZADZIAŁA !!
# Przykład (z pliku Enova):
#   NIP_PLATNIKA = "XXXXXXXXXX"
#   REGON_PLATNIKA = "XXXXXXXXX"
#   PESEL_PLATNIKA = "XXXXXXXXXXX"
#   NAZWA_PLATNIKA = "NAZWA FIRMY"
#   NAZWISKO_PLATNIKA = "NAZWISKO"
#   IMIE_PLATNIKA = "IMIE"
#   DATA_UR_PLATNIKA = "RRRR-MM-DD"

NIP_PLATNIKA = ""              # ← WPISZ NIP FIRMY (obowiązkowe!)
REGON_PLATNIKA = ""            # ← WPISZ REGON
PESEL_PLATNIKA = ""            # ← PESEL właściciela (jeśli osoba fizyczna)
NAZWA_PLATNIKA = ""            # ← Pełna nazwa firmy
NAZWISKO_PLATNIKA = ""         # ← Nazwisko właściciela (osoba fizyczna)
IMIE_PLATNIKA = ""             # ← Imię właściciela
DATA_UR_PLATNIKA = ""          # ← Data urodzenia właściciela (RRRR-MM-DD)
TYP_IDENTYFIKATORA = "P"      # P=PESEL, N=NIP, R=REGON

# === PARAMETRY DEKLARACJI ===
OKRES_ROK = 2026
OKRES_MIESIAC = 3
NR_DEKLARACJI = 1          # Numer kolejny deklaracji w miesiącu (zwykle 1)
STOPA_WYPADKOWA = "1.67"   # Stopa procentowa składki wypadkowej

# === RPA - OPÓŹNIENIA (sekundy) ===
DELAY_SHORT = 0.3           # Krótkie opóźnienie (kliknięcia)
DELAY_MEDIUM = 1.0          # Średnie (ładowanie okien)
DELAY_LONG = 3.0            # Długie (import/eksport)
DELAY_SEND = 5.0            # Wysyłka deklaracji

# === KODY ZUS ===
KODY_TYTULU_UBEZPIECZENIA = {
    "pracownik": "0110",
    "zlecenie_zus": "0411",
    "zlecenie_zdrowotne": "0411",
    "praktyka_absolwencka": "0120",
    "czlonek_rady_nadzorczej": "2241",
}

KODY_NFZ = {
    "dolnoslaski": "01",
    "kujawsko_pomorski": "02",
    "lubelski": "03",
    "lubuski": "04",
    "lodzki": "05",
    "malopolski": "06",
    "mazowiecki": "07",
    "opolski": "08",
    "podkarpacki": "09",
    "podlaski": "10",
    "pomorski": "11",
    "slaski": "12",
    "swietokrzyski": "13",
    "warminsko_mazurski": "14",
    "wielkopolski": "15",
    "zachodniopomorski": "16",
}
