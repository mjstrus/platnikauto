"""
Generator potwierdzeń rejestracji ZUS — osobny PDF per pracownik.
Generuje folder z plikami PDF: jeden pracownik = jeden plik.

Użycie:
  from pdf_confirmations import generuj_potwierdzenia
  sciezki = generuj_potwierdzenia(pracownicy, katalog_wyjsciowy)
"""
import logging
from pathlib import Path
from datetime import datetime
from fpdf import FPDF

import config
from excel_reader import Pracownik

logger = logging.getLogger("platnik_auto")

# Fonty z obsługą polskich znaków — szukaj w kilku lokalizacjach
import os as _os
_FONT_DIRS = [
    "/usr/share/fonts/truetype/dejavu/",          # Linux / Streamlit Cloud
    "C:/Windows/Fonts/",                           # Windows
    "/usr/share/fonts/dejavu/",                    # Niektóre dystrybucje
]
FONT_DIR = ""
for _d in _FONT_DIRS:
    if _os.path.exists(_d + "DejaVuSans.ttf"):
        FONT_DIR = _d
        break
if not FONT_DIR:
    FONT_DIR = _FONT_DIRS[0]  # fallback
FONT_REGULAR = FONT_DIR + "DejaVuSans.ttf"
FONT_BOLD = FONT_DIR + "DejaVuSans-Bold.ttf"

KODY_TYTULU = {
    "0110": "Pracownik — umowa o pracę",
    "0411": "Zleceniobiorca",
    "0120": "Praktyka absolwencka",
    "0510": "Osoba wykonująca pracę na podstawie umowy agencyjnej",
    "2241": "Członek rady nadzorczej",
}

KODY_NFZ = {
    "01": "dolnośląski", "02": "kujawsko-pomorski", "03": "lubelski",
    "04": "lubuski", "05": "łódzki", "06": "małopolski",
    "07": "mazowiecki", "08": "opolski", "09": "podkarpacki",
    "10": "podlaski", "11": "pomorski", "12": "śląski",
    "13": "świętokrzyski", "14": "warmińsko-mazurski",
    "15": "wielkopolski", "16": "zachodniopomorski",
}

TYP_ID_OPIS = {
    "P": "PESEL",
    "1": "Dowód osobisty",
    "2": "Paszport",
}


class PotwierdzeniePDF(FPDF):
    """PDF z potwierdzeniem rejestracji ZUS dla jednego pracownika"""

    def __init__(self):
        super().__init__()
        # Rejestruj fonty z polskimi znakami
        self.add_font("DejaVu", "", FONT_REGULAR, uni=True)
        self.add_font("DejaVu", "B", FONT_BOLD, uni=True)
        self.set_auto_page_break(auto=True, margin=25)

    def header(self):
        # Logo / nagłówek biura
        self.set_font("DejaVu", "B", 11)
        self.set_text_color(31, 78, 121)
        nazwa_biura = getattr(config, 'NAZWA_BIURA', 'ABACUS CENTRUM KSIĘGOWE')
        self.cell(0, 8, nazwa_biura, ln=True, align="L")

        adres_biura = getattr(config, 'ADRES_BIURA', '')
        if adres_biura:
            self.set_font("DejaVu", "", 8)
            self.set_text_color(100, 100, 100)
            self.cell(0, 5, adres_biura, ln=True, align="L")

        # Linia oddzielająca
        self.set_draw_color(31, 78, 121)
        self.set_line_width(0.5)
        self.line(10, self.get_y() + 2, 200, self.get_y() + 2)
        self.ln(8)

    def footer(self):
        self.set_y(-20)
        self.set_draw_color(180, 180, 180)
        self.set_line_width(0.3)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(3)
        self.set_font("DejaVu", "", 7)
        self.set_text_color(150, 150, 150)
        self.cell(0, 4, f"Wygenerowano: {datetime.now().strftime('%Y-%m-%d %H:%M')}  •  Płatnik Auto v2.0", align="C")

    def _tytul(self, tekst):
        self.set_font("DejaVu", "B", 16)
        self.set_text_color(31, 78, 121)
        self.cell(0, 12, tekst, ln=True, align="C")
        self.ln(4)

    def _podtytul(self, tekst):
        self.set_font("DejaVu", "B", 11)
        self.set_text_color(46, 117, 182)
        self.cell(0, 8, tekst, ln=True)
        self.set_draw_color(46, 117, 182)
        self.set_line_width(0.3)
        self.line(10, self.get_y(), 120, self.get_y())
        self.ln(3)

    def _wiersz(self, etykieta, wartosc, bold_value=False):
        self.set_font("DejaVu", "", 10)
        self.set_text_color(100, 100, 100)
        x_start = self.get_x()
        self.cell(65, 7, etykieta, ln=False)
        if bold_value:
            self.set_font("DejaVu", "B", 10)
        self.set_text_color(30, 30, 30)
        self.cell(0, 7, str(wartosc), ln=True)

    def _checkbox(self, etykieta, zaznaczony=True):
        x = self.get_x()
        y = self.get_y()
        # Kwadracik
        self.set_draw_color(100, 100, 100)
        self.set_line_width(0.3)
        self.rect(x, y + 1, 4, 4)
        if zaznaczony:
            # Ptaszek
            self.set_draw_color(16, 185, 129)
            self.set_line_width(0.6)
            self.line(x + 0.8, y + 3, x + 1.8, y + 4.2)
            self.line(x + 1.8, y + 4.2, x + 3.5, y + 1.5)
        self.set_x(x + 7)
        self.set_font("DejaVu", "", 9)
        self.set_text_color(50, 50, 50)
        self.cell(40, 6, etykieta)

    def _odstep(self, h=5):
        self.ln(h)


def generuj_pdf_pracownika(p: Pracownik, sciezka: Path, nr_potw: str = "") -> Path:
    """Generuje pojedynczy PDF potwierdzenia dla pracownika"""
    pdf = PotwierdzeniePDF()
    pdf.add_page()

    # Tytuł
    pdf._tytul("POTWIERDZENIE ZGŁOSZENIA DO ZUS")

    # Typ zgłoszenia
    typ_zgl = p.typ_deklaracji
    if typ_zgl == "ZIUA":
        typ_opis = "ZIUA — Zmiana danych identyfikacyjnych ubezpieczonego"
    elif typ_zgl == "ZZA":
        typ_opis = "ZZA — Zgłoszenie do ubezpieczenia zdrowotnego"
    else:
        typ_opis = "ZUA — Zgłoszenie do ubezpieczeń społecznych i zdrowotnego"

    pdf.set_font("DejaVu", "B", 11)
    pdf.set_text_color(30, 30, 30)
    pdf.cell(0, 8, typ_opis, ln=True, align="C")
    pdf._odstep(6)

    # === DANE PŁATNIKA ===
    pdf._podtytul("Dane płatnika składek")
    pdf._wiersz("Nazwa:", config.NAZWA_PLATNIKA or "—", bold_value=True)
    pdf._wiersz("NIP:", config.NIP_PLATNIKA or "—")
    if config.REGON_PLATNIKA:
        pdf._wiersz("REGON:", config.REGON_PLATNIKA)
    pdf._odstep()

    # === DANE PRACOWNIKA ===
    pdf._podtytul("Dane ubezpieczonego")
    pdf._wiersz("Nazwisko:", p.nazwisko.upper(), bold_value=True)
    pdf._wiersz("Imię:", p.imie.upper(), bold_value=True)
    if p.drugie_imie:
        pdf._wiersz("Drugie imię:", p.drugie_imie.upper())
    if p.data_urodzenia:
        pdf._wiersz("Data urodzenia:", str(p.data_urodzenia))

    # Identyfikator
    if p.pesel:
        pdf._wiersz("PESEL:", p.pesel, bold_value=True)
    if p.nr_paszportu:
        pdf._wiersz("Nr paszportu:", f"{p.seria_paszportu} {p.nr_paszportu}".strip(), bold_value=True)
        if p.kraj_paszportu:
            pdf._wiersz("Kraj wydania:", p.kraj_paszportu)
    if p.obywatelstwo and p.obywatelstwo != "PL":
        pdf._wiersz("Obywatelstwo:", p.obywatelstwo)

    # Adres
    if p.miejscowosc:
        adres = f"{p.ulica} {p.nr_domu}".strip()
        if p.nr_lokalu:
            adres += f"/{p.nr_lokalu}"
        adres += f", {p.kod_pocztowy} {p.miejscowosc}"
        pdf._wiersz("Adres:", adres)
    pdf._odstep()

    # === DANE ZGŁOSZENIA ===
    pdf._podtytul("Dane zgłoszenia")
    kod_tyt = str(p.kod_tytulu).zfill(4)
    pdf._wiersz("Kod tytułu ubezp.:", f"{kod_tyt} — {KODY_TYTULU.get(kod_tyt, '')}")
    kod_nfz = str(p.kod_nfz).zfill(2)
    pdf._wiersz("Oddział NFZ:", f"{kod_nfz} — {KODY_NFZ.get(kod_nfz, '')}")
    if p.data_zgłoszenia:
        pdf._wiersz("Data zgłoszenia:", str(p.data_zgłoszenia))
    if p.data_powstania_obowiazku:
        pdf._wiersz("Data powstania obowiązku:", str(p.data_powstania_obowiazku))
    pdf._odstep()

    # === UBEZPIECZENIA ===
    pdf._podtytul("Rodzaje ubezpieczeń")
    y_start = pdf.get_y()

    # Lewa kolumna
    pdf.set_xy(10, y_start)
    pdf._checkbox("Emerytalne", p.ubezp_emerytalne)
    pdf.set_xy(10, y_start + 7)
    pdf._checkbox("Rentowe", p.ubezp_rentowe)
    pdf.set_xy(10, y_start + 14)
    pdf._checkbox("Chorobowe", p.ubezp_chorobowe)

    # Prawa kolumna
    pdf.set_xy(80, y_start)
    pdf._checkbox("Wypadkowe", p.ubezp_wypadkowe)
    pdf.set_xy(80, y_start + 7)
    pdf._checkbox("Zdrowotne", p.ubezp_zdrowotne)

    pdf.set_y(y_start + 24)
    pdf._odstep()

    # === ZIUA — zmiana danych ===
    if p.czy_zmiana_identyfikatora:
        pdf._podtytul("Zmiana danych identyfikacyjnych (ZIUA)")
        pdf._wiersz("Poprzedni identyfikator:", TYP_ID_OPIS.get(p.poprzedni_typ_id, p.poprzedni_typ_id))
        pdf._wiersz("Poprzedni nr dokumentu:", f"{p.poprzednia_seria} {p.poprzedni_nr_dokumentu}".strip())
        pdf._wiersz("Nowy identyfikator:", "PESEL")
        pdf._wiersz("Nowy PESEL:", p.pesel, bold_value=True)
        pdf._odstep()

    # === POTWIERDZENIE ===
    pdf._podtytul("Potwierdzenie przetwarzania")
    nr = nr_potw or p.nr_potwierdzenia or f"AUTO-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    pdf._wiersz("Nr potwierdzenia:", nr, bold_value=True)
    pdf._wiersz("Data przetwarzania:", p.data_przetwarzania or datetime.now().strftime("%Y-%m-%d"))
    pdf._wiersz("Status:", p.status if p.status != "nowy" else "Zarejestrowano")
    pdf._wiersz("Okres:", f"{config.OKRES_ROK}-{str(config.OKRES_MIESIAC).zfill(2)}")

    # === PODPIS ===
    pdf._odstep(12)
    pdf.set_font("DejaVu", "", 9)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(95, 5, "Podpis ubezpieczonego:", align="C")
    pdf.cell(95, 5, "Podpis i pieczęć płatnika:", align="C", ln=True)
    pdf._odstep(2)
    pdf.set_draw_color(150, 150, 150)
    pdf.set_line_width(0.3)
    # Linie na podpisy
    pdf.line(20, pdf.get_y() + 15, 90, pdf.get_y() + 15)
    pdf.line(115, pdf.get_y() + 15, 190, pdf.get_y() + 15)

    # Zapisz
    sciezka = Path(sciezka)
    sciezka.parent.mkdir(parents=True, exist_ok=True)
    pdf.output(str(sciezka))
    return sciezka


def generuj_potwierdzenia(
    pracownicy: list[Pracownik],
    katalog: str | Path = None,
    nr_potwierdzenia_prefix: str = "",
) -> list[Path]:
    """
    Generuje potwierdzenia PDF dla wszystkich pracowników.
    Zwraca listę ścieżek do wygenerowanych plików.
    
    Args:
        pracownicy: lista pracowników do przetworzenia
        katalog: folder wyjściowy (domyślnie: output/potwierdzenia/)
        nr_potwierdzenia_prefix: prefix nr potwierdzenia (np. "ZUS-2026-03-")
    """
    if katalog is None:
        katalog = Path("output") / "potwierdzenia" / datetime.now().strftime("%Y%m%d_%H%M%S")
    katalog = Path(katalog)
    katalog.mkdir(parents=True, exist_ok=True)

    sciezki = []
    poprawni = [p for p in pracownicy if not p.bledy]
    
    logger.info(f"Generuję potwierdzenia PDF dla {len(poprawni)} pracowników...")

    for idx, p in enumerate(poprawni, 1):
        # Nazwa pliku: TYP_PESEL_NAZWISKO_IMIE.pdf
        id_str = p.pesel or p.nr_paszportu or f"BRAK_{idx}"
        nazwisko_safe = p.nazwisko.upper().replace(" ", "_")
        imie_safe = p.imie.upper().replace(" ", "_")
        typ = p.typ_deklaracji or "ZUA"
        filename = f"{typ}_{id_str}_{nazwisko_safe}_{imie_safe}.pdf"

        # Nr potwierdzenia
        nr_potw = f"{nr_potwierdzenia_prefix}{idx:04d}" if nr_potwierdzenia_prefix else ""

        sciezka = katalog / filename
        try:
            generuj_pdf_pracownika(p, sciezka, nr_potw)
            sciezki.append(sciezka)
            if idx % 50 == 0:
                logger.info(f"  Wygenerowano {idx}/{len(poprawni)}...")
        except Exception as e:
            logger.error(f"  Błąd PDF dla {p.nazwisko} {p.imie}: {e}")

    logger.info(f"Gotowe: {len(sciezki)} potwierdzeń w {katalog}")
    return sciezki


# === CLI ===
if __name__ == "__main__":
    """Test — generuje potwierdzenia z pliku Excel"""
    import sys
    from excel_reader import wczytaj_excel

    if len(sys.argv) > 1:
        plik = sys.argv[1]
    else:
        plik = "dane_pracownicy.xlsx"

    try:
        pracownicy = wczytaj_excel(plik)
        sciezki = generuj_potwierdzenia(pracownicy, nr_potwierdzenia_prefix="ZUS-2026-03-")
        print(f"\nWygenerowano {len(sciezki)} potwierdzeń:")
        for s in sciezki:
            print(f"  {s}")
    except FileNotFoundError:
        print(f"Plik nie znaleziony: {plik}")
