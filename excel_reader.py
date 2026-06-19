"""
Moduł odczytu i walidacji danych z Excela
Obsługuje dane pracowników dla deklaracji RCA, DRA, ZUA/ZZA
"""
import pandas as pd
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional
import re
import logging

logger = logging.getLogger("platnik_auto")


@dataclass
class Pracownik:
    """Dane pracownika potrzebne do deklaracji ZUS"""
    # Identyfikacja
    pesel: str = ""
    nazwisko: str = ""
    imie: str = ""
    drugie_imie: str = ""
    data_urodzenia: str = ""
    
    # Adres
    kod_pocztowy: str = ""
    miejscowosc: str = ""
    ulica: str = ""
    nr_domu: str = ""
    nr_lokalu: str = ""
    
    # ZUS
    kod_tytulu: str = "0110"  # domyślnie: pracownik
    kod_nfz: str = ""         # auto-wykrywany z kodu pocztowego
    
    # Wynagrodzenie / składki (RCA)
    podstawa_emerytalna: float = 0.0
    podstawa_rentowa: float = 0.0
    podstawa_chorobowa: float = 0.0
    podstawa_wypadkowa: float = 0.0
    podstawa_zdrowotna: float = 0.0
    podstawa_fp_fgsp: float = 0.0
    
    # Składki obliczone
    skladka_emerytalna_pracownik: float = 0.0
    skladka_emerytalna_pracodawca: float = 0.0
    skladka_rentowa_pracownik: float = 0.0
    skladka_rentowa_pracodawca: float = 0.0
    skladka_chorobowa: float = 0.0
    skladka_wypadkowa: float = 0.0
    skladka_zdrowotna: float = 0.0
    skladka_fp: float = 0.0
    skladka_fgsp: float = 0.0
    
    # ZUA/ZZA
    data_zgłoszenia: str = ""
    data_powstania_obowiazku: str = ""
    typ_deklaracji: str = "RCA"  # RCA, ZUA, ZZA, ZIUA
    
    # Obcokrajowcy / identyfikator
    typ_identyfikatora: str = "P"   # P=PESEL, 1=dowód, 2=paszport
    nr_paszportu: str = ""
    seria_paszportu: str = ""
    kraj_paszportu: str = ""         # kod kraju np. "UA", "BY"
    obywatelstwo: str = "PL"
    czy_obcokrajowiec: bool = False
    
    # ZIUA — zmiana identyfikatora (paszport → PESEL)
    poprzedni_typ_id: str = ""       # np. "2" (paszport)
    poprzedni_nr_dokumentu: str = "" # stary nr paszportu
    poprzednia_seria: str = ""
    czy_zmiana_identyfikatora: bool = False  # True → generuj ZIUA
    
    # Ubezpieczenia (do potwierdzeń)
    ubezp_emerytalne: bool = True
    ubezp_rentowe: bool = True
    ubezp_chorobowe: bool = True
    ubezp_wypadkowe: bool = True
    ubezp_zdrowotne: bool = True
    
    # Status przetwarzania
    status: str = "nowy"
    nr_potwierdzenia: str = ""
    data_przetwarzania: str = ""
    bledy: list = field(default_factory=list)


def waliduj_pesel(pesel: str) -> bool:
    if not pesel or len(pesel) != 11 or not pesel.isdigit():
        return False
    wagi = [1, 3, 7, 9, 1, 3, 7, 9, 1, 3]
    suma = sum(int(p) * w for p, w in zip(pesel, wagi))
    return (10 - suma % 10) % 10 == int(pesel[10])


def waliduj_kod_pocztowy(kod: str) -> bool:
    return bool(re.match(r'^\d{2}-\d{3}$', str(kod).strip()))


# === AUTO-DETEKCJA NFZ Z KODU POCZTOWEGO ===
_KOD_NFZ_MAP = {
    range(0,10): "07",   # mazowiecki
    range(10,15): "14",  # warmińsko-mazurski
    range(15,20): "10",  # podlaski
    range(20,25): "03",  # lubelski
    range(25,30): "13",  # świętokrzyski
    range(30,35): "06",  # małopolski
    range(35,40): "09",  # podkarpacki
    range(40,45): "12",  # śląski
    range(45,50): "08",  # opolski
    range(50,60): "01",  # dolnośląski
    range(60,65): "15",  # wielkopolski
    range(65,70): "04",  # lubuski
    range(70,80): "16",  # zachodniopomorski
    range(80,85): "11",  # pomorski
    range(85,90): "02",  # kujawsko-pomorski
    range(90,100): "05", # łódzki
}

def kod_pocztowy_do_nfz(kod_pocztowy: str) -> str:
    """Wykrywa oddział NFZ na podstawie kodu pocztowego"""
    if not kod_pocztowy:
        return ""
    try:
        prefix = int(str(kod_pocztowy).replace("-", "").strip()[:2])
        for r, nfz in _KOD_NFZ_MAP.items():
            if prefix in r:
                return nfz
    except (ValueError, IndexError):
        pass
    return ""


def waliduj_pracownika(p: Pracownik) -> list[str]:
    bledy = []
    # Identyfikacja: PESEL lub paszport
    if p.typ_identyfikatora == "P" or (p.pesel and not p.nr_paszportu):
        if not waliduj_pesel(p.pesel):
            bledy.append(f"Nieprawidłowy PESEL: {p.pesel}")
    elif p.typ_identyfikatora == "2" or p.nr_paszportu:
        if not p.nr_paszportu.strip():
            bledy.append("Brak numeru paszportu dla obcokrajowca")
    else:
        bledy.append("Brak identyfikatora (PESEL lub paszport)")
    
    if not p.nazwisko.strip():
        bledy.append("Brak nazwiska")
    if not p.imie.strip():
        bledy.append("Brak imienia")
    if p.typ_deklaracji == "RCA":
        if p.podstawa_emerytalna <= 0 and p.podstawa_zdrowotna <= 0:
            bledy.append("Brak podstaw składkowych dla RCA")
    if p.kod_pocztowy and not waliduj_kod_pocztowy(p.kod_pocztowy):
        bledy.append(f"Nieprawidłowy kod pocztowy: {p.kod_pocztowy}")
    
    # ZIUA: walidacja danych do zmiany
    if p.czy_zmiana_identyfikatora:
        if not p.poprzedni_nr_dokumentu.strip():
            bledy.append("ZIUA: brak poprzedniego numeru dokumentu")
        if not p.pesel.strip():
            bledy.append("ZIUA: brak nowego PESEL")
    
    # Auto-detekcja obcokrajowca
    if p.nr_paszportu and not p.czy_obcokrajowiec:
        p.czy_obcokrajowiec = True
    if p.obywatelstwo and p.obywatelstwo != "PL":
        p.czy_obcokrajowiec = True
    
    # Auto-detekcja ZIUA
    if p.pesel and p.poprzedni_nr_dokumentu and not p.czy_zmiana_identyfikatora:
        p.czy_zmiana_identyfikatora = True
        p.typ_deklaracji = "ZIUA"
    
    return bledy


# === MAPOWANIE KOLUMN EXCEL → PRACOWNIK ===
# Format 1: Szablon rozszerzony (generuj_szablon.py)
KOLUMNY_MAP = {
    "PESEL": "pesel",
    "Nazwisko": "nazwisko",
    "Imię": "imie",
    "Drugie imię": "drugie_imie",
    "Data urodzenia": "data_urodzenia",
    "Kod pocztowy": "kod_pocztowy",
    "Miejscowość": "miejscowosc",
    "Ulica": "ulica",
    "Nr domu": "nr_domu",
    "Nr lokalu": "nr_lokalu",
    "Kod tytułu": "kod_tytulu",
    "Kod NFZ": "kod_nfz",
    "Podstawa emerytalna": "podstawa_emerytalna",
    "Podstawa rentowa": "podstawa_rentowa",
    "Podstawa chorobowa": "podstawa_chorobowa",
    "Podstawa wypadkowa": "podstawa_wypadkowa",
    "Podstawa zdrowotna": "podstawa_zdrowotna",
    "Podstawa FP/FGŚP": "podstawa_fp_fgsp",
    "Składka emer. pracownik": "skladka_emerytalna_pracownik",
    "Składka emer. pracodawca": "skladka_emerytalna_pracodawca",
    "Składka rent. pracownik": "skladka_rentowa_pracownik",
    "Składka rent. pracodawca": "skladka_rentowa_pracodawca",
    "Składka chorobowa": "skladka_chorobowa",
    "Składka wypadkowa": "skladka_wypadkowa",
    "Składka zdrowotna": "skladka_zdrowotna",
    "Składka FP": "skladka_fp",
    "Składka FGŚP": "skladka_fgsp",
    "Data zgłoszenia": "data_zgłoszenia",
    "Data powstania obowiązku": "data_powstania_obowiazku",
    "Typ deklaracji": "typ_deklaracji",
    "Typ identyfikatora": "typ_identyfikatora",
    "Nr paszportu": "nr_paszportu",
    "Seria paszportu": "seria_paszportu",
    "Kraj paszportu": "kraj_paszportu",
    "Obywatelstwo": "obywatelstwo",
    "Poprzedni typ ID": "poprzedni_typ_id",
    "Poprzedni nr dokumentu": "poprzedni_nr_dokumentu",
    "Poprzednia seria": "poprzednia_seria",
}

# Format 2: Format biurowy (plik od klienta)
KOLUMNY_MAP_BIURO = {
    "PESEL": "pesel",
    "Nazwisko": "nazwisko",
    "Imię": "imie",
    "drugie imię": "drugie_imie",
    "Imię i nazwisko": "_imie_nazwisko",   # specjalne: rozdzielimy
    "Data urodzenia": "data_urodzenia",
    "Rozpoczęcie pracy ZUS": "data_zgłoszenia",
    "Zakończenie pracy ZUS": "_zakonczenie",
    "Typ dokumentu": "_typ_dokumentu",      # paszport/dowód → typ_identyfikatora
    "Nr dokumentu": "nr_paszportu",
    "Obywatelstwo": "obywatelstwo",
    "Płeć": "_plec",
    "Rodzaj umowy": "_rodzaj_umowy",
    "Oddział NFZ": "kod_nfz",
    "Kod Pocztowy": "kod_pocztowy",
    "Miejscowość": "miejscowosc",
    "Ulica": "ulica",
    "Numer domu": "nr_domu",
    "Numer lokalu": "nr_lokalu",
    "ZUS": "_zus",
    "Stanowisko": "_stanowisko",
    "Kod tytułu": "kod_tytulu",
    "Kod zgłoszenia": "typ_deklaracji",
}


def _wykryj_format(kolumny_excel: list[str]) -> dict:
    """Auto-detekcja formatu Excel na podstawie nagłówków"""
    kolumny_norm = [str(k).strip() for k in kolumny_excel]
    
    if any(k in kolumny_norm for k in ["Imię i nazwisko", "Typ dokumentu", "Rozpoczęcie pracy ZUS", "Numer domu"]):
        logger.info("Wykryto format: biurowy (plik od klienta)")
        return KOLUMNY_MAP_BIURO
    else:
        logger.info("Wykryto format: szablon rozszerzony")
        return KOLUMNY_MAP


def wczytaj_excel(sciezka: str | Path, arkusz: str = None) -> list[Pracownik]:
    """
    Wczytuje dane pracowników z pliku Excel.
    Auto-wykrywa format (biurowy vs szablon rozszerzony).
    Zwraca listę obiektów Pracownik z walidacją.
    """
    sciezka = Path(sciezka)
    if not sciezka.exists():
        raise FileNotFoundError(f"Nie znaleziono pliku: {sciezka}")

    df = pd.read_excel(sciezka, sheet_name=arkusz or 0, dtype={
        "PESEL": str, "Kod pocztowy": str, "Kod Pocztowy": str,
        "Kod tytułu": str, "Kod NFZ": str, "Oddział NFZ": str,
        "Nr dokumentu": str, "Nr paszportu": str,
    })
    logger.info(f"Wczytano {len(df)} wierszy z {sciezka.name}")

    df.columns = df.columns.str.strip()
    
    # Auto-detekcja formatu
    mapa = _wykryj_format(list(df.columns))
    is_biuro = (mapa is KOLUMNY_MAP_BIURO)

    pracownicy = []
    for idx, row in df.iterrows():
        kwargs = {}
        extra = {}  # pola specjalne zaczynające się od _
        
        for excel_col, attr in mapa.items():
            if excel_col in df.columns:
                val = row[excel_col]
                if pd.isna(val):
                    continue
                if isinstance(val, float) and attr in ("pesel", "kod_pocztowy", "kod_tytulu", "kod_nfz", "nr_domu", "nr_lokalu", "nr_paszportu"):
                    val = str(int(val))
                    if attr == "kod_tytulu":
                        val = val.zfill(4)
                    elif attr == "kod_nfz":
                        val = val.zfill(2)
                else:
                    val = val if isinstance(val, (int, float)) else str(val).strip()
                
                if attr.startswith("_"):
                    extra[attr] = val
                else:
                    kwargs[attr] = val

        # === Konwersje dla formatu biurowego ===
        if is_biuro:
            # Typ dokumentu → typ_identyfikatora + nr_paszportu
            typ_dok = str(extra.get("_typ_dokumentu", "")).lower()
            if "paszport" in typ_dok:
                kwargs["typ_identyfikatora"] = "2"
                kwargs.setdefault("czy_obcokrajowiec", True)
            elif "dowód" in typ_dok or "dowod" in typ_dok:
                kwargs["typ_identyfikatora"] = "1"
            
            # Obywatelstwo → czy_obcokrajowiec
            obyw = str(kwargs.get("obywatelstwo", "")).strip()
            if obyw and obyw.upper() not in ("PL", "POLSKA", "POLAND", ""):
                kwargs["czy_obcokrajowiec"] = True
            
            # Imię i nazwisko → rozdziel jeśli brak osobnych pól
            if not kwargs.get("nazwisko") and extra.get("_imie_nazwisko"):
                parts = str(extra["_imie_nazwisko"]).strip().split()
                if len(parts) >= 2:
                    kwargs["imie"] = parts[0]
                    kwargs["nazwisko"] = " ".join(parts[1:])
                elif len(parts) == 1:
                    # Jedno imię (np. indonezyjskie) — użyj jako nazwisko
                    kwargs["nazwisko"] = parts[0]
                    kwargs.setdefault("imie", parts[0])
            
            # Jeśli imię jest wypełnione ale nazwisko nie — jednoimienni
            if kwargs.get("imie") and not kwargs.get("nazwisko"):
                kwargs["nazwisko"] = kwargs["imie"]
            
            # Data zgłoszenia z "Rozpoczęcie pracy ZUS"
            if kwargs.get("data_zgłoszenia"):
                kwargs["data_powstania_obowiazku"] = kwargs["data_zgłoszenia"]
            
            # Domyślny typ deklaracji: ZUA (zgłoszenie)
            if "typ_deklaracji" not in kwargs:
                kwargs["typ_deklaracji"] = "ZUA"

        try:
            # Pomiń puste wiersze
            if not kwargs.get("pesel") and not kwargs.get("nazwisko") and not kwargs.get("imie") and not kwargs.get("nr_paszportu"):
                continue
            
            # PESEL: dopełnij do 11 cyfr zerem (osoby urodzone po 2000)
            pesel_val = str(kwargs.get("pesel", "")).strip()
            if pesel_val and pesel_val.isdigit() and len(pesel_val) == 10:
                kwargs["pesel"] = "0" + pesel_val
                logger.info(f"Auto-dopełnienie PESEL: {pesel_val} → 0{pesel_val}")
            
            p = Pracownik(**kwargs)
            
            # Auto-detekcja NFZ z kodu pocztowego
            if not p.kod_nfz and p.kod_pocztowy:
                p.kod_nfz = kod_pocztowy_do_nfz(p.kod_pocztowy)
                if p.kod_nfz:
                    logger.debug(f"Auto NFZ: {p.kod_pocztowy} → {p.kod_nfz}")
            
            p.bledy = waliduj_pracownika(p)
            if p.bledy:
                p.status = "błąd_walidacji"
                logger.warning(f"Wiersz {idx+2}: {p.nazwisko} {p.imie} — {', '.join(p.bledy)}")
            pracownicy.append(p)
        except Exception as e:
            logger.error(f"Wiersz {idx+2}: błąd tworzenia rekordu — {e}")

    poprawne = sum(1 for p in pracownicy if p.status == "nowy")
    bledne = sum(1 for p in pracownicy if p.status == "błąd_walidacji")
    logger.info(f"Wynik: {poprawne} poprawnych, {bledne} z błędami, {len(pracownicy)} łącznie")
    return pracownicy


def filtruj_wg_typu(pracownicy: list[Pracownik], typ: str) -> list[Pracownik]:
    return [p for p in pracownicy if p.typ_deklaracji.upper() == typ.upper()]


def raport_walidacji(pracownicy: list[Pracownik]) -> str:
    lines = ["=" * 60, "RAPORT WALIDACJI DANYCH", "=" * 60, ""]
    
    poprawne = [p for p in pracownicy if not p.bledy]
    bledne = [p for p in pracownicy if p.bledy]
    
    lines.append(f"Łącznie rekordów:  {len(pracownicy)}")
    lines.append(f"Poprawnych:        {len(poprawne)}")
    lines.append(f"Z błędami:         {len(bledne)}")
    lines.append("")
    
    if bledne:
        lines.append("--- BŁĘDY ---")
        for p in bledne:
            lines.append(f"  {p.pesel} | {p.nazwisko} {p.imie}")
            for b in p.bledy:
                lines.append(f"    ✗ {b}")
        lines.append("")
    
    # Podsumowanie typów deklaracji
    typy = {}
    for p in poprawne:
        typy[p.typ_deklaracji] = typy.get(p.typ_deklaracji, 0) + 1
    lines.append("--- WG TYPU DEKLARACJI ---")
    for typ, cnt in sorted(typy.items()):
        lines.append(f"  {typ}: {cnt} pracowników")
    
    return "\n".join(lines)
