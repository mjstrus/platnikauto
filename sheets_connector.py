"""
Integracja z Google Sheets — zapis danych ze skanera dokumentów.
Używa gspread + service account.

Setup:
  1. Utwórz Service Account w Google Cloud Console
  2. Pobierz klucz JSON
  3. Udostępnij arkusz dla emaila service account
  4. Wklej klucz JSON w Streamlit secrets lub podaj ścieżkę
"""
import json
import logging
from datetime import datetime

logger = logging.getLogger("platnik_auto")

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False


# Mapowanie kolumn arkusza (na podstawie struktury "import xyz")
# Kolumna H=8, J=10, K=11, L=12, M=13, N=14, O=15, P=16, G=7
KOLUMNY_SHEET = {
    "name_surname": "H",       # Name and Surname
    "contract": "I",            # Contract
    "passport_number": "J",     # Number passport
    "passport_issued": "K",     # Date issued
    "passport_expiry": "L",     # Date of expiry
    "citizenship": "M",        # Citizenship
    "pesel": "N",              # PESEL
    "birth": "O",              # BIRTH
    "gender": "P",             # Gender
    "zgloszenie_zus": "G",     # zgłoś do ZUS
}

# Numer kolumny (1-indexed) z litery
def _col_num(letter):
    result = 0
    for c in letter.upper():
        result = result * 26 + (ord(c) - ord('A') + 1)
    return result


class SheetsConnector:
    """Łączy się z Google Sheets i zapisuje dane pracowników"""
    
    def __init__(self, credentials_json: str | dict = None, credentials_path: str = None):
        if not GSPREAD_AVAILABLE:
            raise RuntimeError("Zainstaluj: pip install gspread google-auth")
        
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        
        if credentials_json:
            if isinstance(credentials_json, str):
                credentials_json = json.loads(credentials_json)
            creds = Credentials.from_service_account_info(credentials_json, scopes=scopes)
        elif credentials_path:
            creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
        else:
            raise ValueError("Podaj credentials_json lub credentials_path")
        
        self.client = gspread.authorize(creds)
        logger.info("Połączono z Google Sheets")
    
    def otworz_arkusz(self, spreadsheet_id: str, sheet_name: str = None, gid: int = None):
        """Otwiera arkusz po ID"""
        self.spreadsheet = self.client.open_by_key(spreadsheet_id)
        
        if gid is not None:
            # Znajdź arkusz po GID
            for ws in self.spreadsheet.worksheets():
                if ws.id == gid:
                    self.worksheet = ws
                    break
            else:
                self.worksheet = self.spreadsheet.sheet1
        elif sheet_name:
            self.worksheet = self.spreadsheet.worksheet(sheet_name)
        else:
            self.worksheet = self.spreadsheet.sheet1
        
        logger.info(f"Otwarto arkusz: {self.spreadsheet.title} / {self.worksheet.title}")
        return self.worksheet
    
    def znajdz_wolny_wiersz(self) -> int:
        """Znajduje pierwszy wolny wiersz (po ostatnim z danymi)"""
        values = self.worksheet.col_values(_col_num("H"))  # kolumna Name and Surname
        return len(values) + 1
    
    def zapisz_pracownika(self, dane: dict, wiersz: int = None) -> int:
        """
        Zapisuje dane pracownika do arkusza.
        
        dane = {
            "name_surname": "NOGUERA SOLARTE MAGALY",
            "passport_number": "BD113815",
            "passport_issued": "27.10.2023",
            "passport_expiry": "26.10.2033",
            "citizenship": "Kolumbia",
            "pesel": "86062421904",
            "birth": "24.06.1986",
            "gender": "K",
        }
        """
        if wiersz is None:
            wiersz = self.znajdz_wolny_wiersz()
        
        cells_to_update = []
        for pole, kolumna_litera in KOLUMNY_SHEET.items():
            if pole in dane and dane[pole]:
                col_num = _col_num(kolumna_litera)
                cells_to_update.append(
                    gspread.Cell(wiersz, col_num, str(dane[pole]))
                )
        
        if cells_to_update:
            self.worksheet.update_cells(cells_to_update)
            logger.info(f"Zapisano wiersz {wiersz}: {dane.get('name_surname', '?')}")
        
        return wiersz
    
    def zapisz_z_dokumentu(self, dane_dokumentu, wiersz: int = None) -> int:
        """
        Zapisuje dane z DaneZDokumentu (ze skanera) do arkusza.
        """
        from document_scanner import DaneZDokumentu
        d = dane_dokumentu
        
        # Złóż imię i nazwisko
        name = f"{d.nazwisko} {d.imie}".strip()
        if d.drugie_imie:
            name = f"{d.nazwisko} {d.imie} {d.drugie_imie}".strip()
        
        dane = {
            "name_surname": name,
            "passport_number": d.nr_dokumentu,
            "passport_issued": "",  # skaner może to odczytać z paszportu
            "passport_expiry": d.data_waznosci,
            "citizenship": d.obywatelstwo or d.kraj_wydania,
            "pesel": d.pesel,
            "birth": _format_date_pl(d.data_urodzenia),
            "gender": d.plec,
            "zgloszenie_zus": "TAK",
        }
        
        return self.zapisz_pracownika(dane, wiersz)
    
    def zapisz_wielu(self, lista_danych: list[dict]) -> list[int]:
        """Zapisuje wielu pracowników naraz"""
        wiersze = []
        start = self.znajdz_wolny_wiersz()
        
        for i, dane in enumerate(lista_danych):
            wiersz = self.zapisz_pracownika(dane, start + i)
            wiersze.append(wiersz)
        
        logger.info(f"Zapisano {len(wiersze)} pracowników (wiersze {start}-{start+len(wiersze)-1})")
        return wiersze


def _format_date_pl(date_str: str) -> str:
    """Konwertuje YYYY-MM-DD na DD.MM.YYYY"""
    if not date_str:
        return ""
    try:
        if "-" in date_str:
            parts = date_str.split("-")
            if len(parts[0]) == 4:  # YYYY-MM-DD
                return f"{parts[2]}.{parts[1]}.{parts[0]}"
    except Exception:
        pass
    return date_str
