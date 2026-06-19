"""
Generator plików KEDU XML — format zgodny z Płatnikiem / KEDU_5_7.
Odwzorowany 1:1 z plików eksportowanych przez Enova365.
Obsługuje: ZUSDRA, ZUSRCA, ZUSZUA.

Namespace: http://www.zus.pl/2026/KEDU_5_7
"""
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime
import logging
from excel_reader import Pracownik
import config

logger = logging.getLogger("platnik_auto")

KEDU_NS = "http://www.zus.pl/2026/KEDU_5_7"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"
KEDU_XSD = "kedu_5_7.xsd"


def _q(kwota) -> str:
    return f"{float(kwota):.2f}"


def _el(parent, tag, text=None):
    e = ET.SubElement(parent, tag)
    if text is not None:
        e.text = str(text)
    return e


def _buduj_naglowek(root):
    nag = _el(root, "naglowek.KEDU")
    prog = _el(nag, "program")
    _el(prog, "producent", "PlatnikAuto")
    _el(prog, "symbol", "PlatnikAuto")
    _el(prog, "wersja", "2.0")
    _el(nag, "data_utworzenia_KEDU", datetime.now().strftime("%Y-%m-%d"))


def _buduj_platnik(parent):
    sek = _el(parent, "II")
    nip = str(config.NIP_PLATNIKA).strip()
    if not nip:
        raise ValueError("BŁĄD: NIP_PLATNIKA jest pusty w config.py! Uzupełnij dane firmy.")
    _el(sek, "p1", nip)
    regon = str(config.REGON_PLATNIKA).strip()
    if regon:
        _el(sek, "p2", regon)
    pesel = str(config.PESEL_PLATNIKA).strip()
    if pesel:
        _el(sek, "p3", pesel)
    nazwa = str(config.NAZWA_PLATNIKA).strip()
    nazwa_skr = str(getattr(config, 'NAZWA_SKROCONA', '')).strip()
    if nazwa_skr:
        _el(sek, "p6", nazwa_skr[:31])
    elif nazwa:
        _el(sek, "p6", nazwa[:31])
    if hasattr(config, 'NAZWISKO_PLATNIKA') and str(config.NAZWISKO_PLATNIKA).strip():
        _el(sek, "p7", str(config.NAZWISKO_PLATNIKA).strip())
    if hasattr(config, 'IMIE_PLATNIKA') and str(config.IMIE_PLATNIKA).strip():
        _el(sek, "p8", str(config.IMIE_PLATNIKA).strip())
    if hasattr(config, 'DATA_UR_PLATNIKA') and str(config.DATA_UR_PLATNIKA).strip():
        _el(sek, "p9", str(config.DATA_UR_PLATNIKA).strip())


def _okres_str():
    return f"{config.OKRES_ROK}-{str(config.OKRES_MIESIAC).zfill(2)}"


# ============================================================
#  ZUSRCA
# ============================================================
def _buduj_zusrca(root, pracownicy, doc_id=2):
    rca = ET.SubElement(root, "ZUSRCA")
    rca.set("id_dokumentu", str(doc_id))

    sek_i = _el(rca, "I")
    p1 = _el(sek_i, "p1")
    _el(p1, "p1", f"{config.NR_DEKLARACJI:02d}")
    _el(p1, "p2", _okres_str())

    _buduj_platnik(rca)

    poprawni = [p for p in pracownicy if not p.bledy]
    for idx, p in enumerate(poprawni, 1):
        blok = ET.SubElement(rca, "III")
        blok.set("id_bloku", f" {idx} ")

        # A — identyfikacja ubezpieczonego
        a = _el(blok, "A")
        _el(a, "p1", p.nazwisko.upper())
        _el(a, "p2", p.imie.upper())
        _el(a, "p3", "P")
        _el(a, "p4", str(p.pesel))

        # B — składki społeczne
        b = _el(blok, "B")
        p1b = _el(b, "p1")
        _el(p1b, "p1", str(p.kod_tytulu).zfill(4))
        _el(p1b, "p2", "0")
        _el(p1b, "p3", "0")

        if str(p.kod_tytulu).zfill(4) == "0110":
            p3b = _el(b, "p3")
            _el(p3b, "p1", "001")
            _el(p3b, "p2", "002")

        _el(b, "p4", _q(p.podstawa_emerytalna))
        _el(b, "p5", _q(p.podstawa_rentowa))
        _el(b, "p6", _q(p.podstawa_chorobowa or p.podstawa_emerytalna))
        _el(b, "p7", _q(p.skladka_emerytalna_pracownik + p.skladka_emerytalna_pracodawca))
        _el(b, "p8", _q(p.skladka_rentowa_pracownik + p.skladka_rentowa_pracodawca))
        _el(b, "p9", _q(p.skladka_chorobowa))
        _el(b, "p10", _q(p.skladka_wypadkowa))
        _el(b, "p11", _q(p.skladka_emerytalna_pracownik))
        _el(b, "p12", _q(p.skladka_rentowa_pracownik))
        _el(b, "p13", "0.00")
        _el(b, "p14", _q(p.skladka_wypadkowa))
        for t in ("p15","p16","p17","p18","p19","p21","p22","p23","p24","p26"):
            _el(b, t, "0.00")
        suma = (p.skladka_emerytalna_pracownik + p.skladka_emerytalna_pracodawca +
                p.skladka_rentowa_pracownik + p.skladka_rentowa_pracodawca +
                p.skladka_chorobowa + p.skladka_wypadkowa)
        _el(b, "p29", _q(suma))

        # C — zdrowotna
        c = _el(blok, "C")
        _el(c, "p1", _q(p.podstawa_zdrowotna))
        _el(c, "p2", "0.00")
        _el(c, "p3", "0.00")
        _el(c, "p4", _q(p.skladka_zdrowotna))
        _el(c, "p5", "0.00")

        # D
        d = _el(blok, "D")
        _el(d, "p1", "0.00")
        _el(d, "p2", "0.00")
        _el(d, "p3", "0.00")

        _el(blok, "E")
        _el(blok, "F")

    sek_iv = _el(rca, "IV")
    _el(sek_iv, "p1", datetime.now().strftime("%Y-%m-%d"))
    return rca


# ============================================================
#  ZUSDRA
# ============================================================
def _buduj_zusdra(root, pracownicy, doc_id=1):
    dra = ET.SubElement(root, "ZUSDRA")
    dra.set("id_dokumentu", str(doc_id))
    poprawni = [p for p in pracownicy if not p.bledy]

    sek_i = _el(dra, "I")
    _el(sek_i, "p1", "6")
    p2i = _el(sek_i, "p2")
    _el(p2i, "p1", f"{config.NR_DEKLARACJI:02d}")
    _el(p2i, "p2", _okres_str())

    _buduj_platnik(dra)

    sek_iii = _el(dra, "III")
    _el(sek_iii, "p1", str(len(poprawni)))
    _el(sek_iii, "p3", str(getattr(config, 'STOPA_WYPADKOWA', '1.67')))

    # IV — sumy składek społecznych
    sek_iv = _el(dra, "IV")
    s_em_pr = sum(p.skladka_emerytalna_pracownik for p in poprawni)
    s_em_pd = sum(p.skladka_emerytalna_pracodawca for p in poprawni)
    s_rt_pr = sum(p.skladka_rentowa_pracownik for p in poprawni)
    s_rt_pd = sum(p.skladka_rentowa_pracodawca for p in poprawni)
    s_ch = sum(p.skladka_chorobowa for p in poprawni)
    s_wp = sum(p.skladka_wypadkowa for p in poprawni)
    s_em = s_em_pr + s_em_pd
    s_rt = s_rt_pr + s_rt_pd
    s_sp = s_em + s_rt + s_ch + s_wp

    _el(sek_iv,"p1",_q(s_em_pr)); _el(sek_iv,"p2",_q(s_em_pd)); _el(sek_iv,"p3",_q(s_em))
    _el(sek_iv,"p4",_q(s_rt_pr)); _el(sek_iv,"p5",_q(s_rt_pd)); _el(sek_iv,"p6",_q(s_rt))
    _el(sek_iv,"p7",_q(s_ch)); _el(sek_iv,"p8",_q(s_wp)); _el(sek_iv,"p9",_q(s_sp))
    for t in ("p10","p11","p12","p13","p14","p15","p16","p17","p18"):
        _el(sek_iv, t, "0.00")
    _el(sek_iv,"p19",_q(s_em_pr)); _el(sek_iv,"p20",_q(s_em_pd)); _el(sek_iv,"p21",_q(s_em))
    _el(sek_iv,"p22",_q(s_rt_pr)); _el(sek_iv,"p23",_q(s_rt_pd)); _el(sek_iv,"p24",_q(s_rt))
    _el(sek_iv,"p25","0.00"); _el(sek_iv,"p26",_q(s_ch)); _el(sek_iv,"p27",_q(s_ch))
    for t in ("p28","p29","p30","p31","p32","p33","p34","p35","p36"):
        _el(sek_iv, t, "0.00")
    _el(sek_iv, "p37", _q(s_sp))

    # V
    sek_v = _el(dra, "V")
    for t in ("p1","p2","p3","p4","p5"):
        _el(sek_v, t, "0.00")

    # VI — zdrowotna
    sek_vi = _el(dra, "VI")
    s_zd = sum(p.skladka_zdrowotna for p in poprawni)
    _el(sek_vi,"p1","0.00"); _el(sek_vi,"p2",_q(s_zd)); _el(sek_vi,"p3","0.00")
    _el(sek_vi,"p4","0.00"); _el(sek_vi,"p5",_q(s_zd)); _el(sek_vi,"p6","0.00")
    _el(sek_vi,"p7",_q(s_zd))

    # VII — FP/FGŚP
    sek_vii = _el(dra, "VII")
    s_fp = sum(p.skladka_fp for p in poprawni)
    s_fg = sum(p.skladka_fgsp for p in poprawni)
    _el(sek_vii,"p1",_q(s_fp)); _el(sek_vii,"p2",_q(s_fg)); _el(sek_vii,"p3",_q(s_fp+s_fg))

    _el(dra, "VIII")

    sek_ix = _el(dra, "IX")
    _el(sek_ix, "p2", _q(s_sp + s_zd + s_fp + s_fg))

    _el(dra, "X")
    _el(dra, "XI")
    _el(dra, "XII")

    sek_xiii = _el(dra, "XIII")
    _el(sek_xiii, "p1", datetime.now().strftime("%Y-%m-%d"))
    return dra


# ============================================================
#  HELPERS
# ============================================================

# Odmiana obywatelstwa: kraj → przymiotnik (rodzaj nijaki)
_OBYWATELSTWO_MAP = {
    # Kraje najczęściej spotykane w biurach rachunkowych
    "POLSKA": "POLSKIE", "PL": "POLSKIE", "POLAND": "POLSKIE",
    "UKRAINA": "UKRAIŃSKIE", "UA": "UKRAIŃSKIE", "UKRAINE": "UKRAIŃSKIE",
    "BIAŁORUŚ": "BIAŁORUSKIE", "BY": "BIAŁORUSKIE", "BELARUS": "BIAŁORUSKIE",
    "INDIE": "INDYJSKIE", "INDIA": "INDYJSKIE", "IN": "INDYJSKIE",
    "INDONEZJA": "INDONEZYJSKIE", "INDONESIA": "INDONEZYJSKIE", "ID": "INDONEZYJSKIE",
    "FILIPINY": "FILIPIŃSKIE", "PHILIPPINES": "FILIPIŃSKIE", "PH": "FILIPIŃSKIE",
    "KOLUMBIA": "KOLUMBIJSKIE", "COLOMBIA": "KOLUMBIJSKIE", "CO": "KOLUMBIJSKIE",
    "NEPAL": "NEPALSKIE", "NP": "NEPALSKIE",
    "BANGLADESZ": "BANGLADESKIE", "BANGLADESH": "BANGLADESKIE", "BD": "BANGLADESKIE",
    "WIETNAM": "WIETNAMSKIE", "VIETNAM": "WIETNAMSKIE", "VN": "WIETNAMSKIE",
    "GRUZJA": "GRUZIŃSKIE", "GEORGIA": "GRUZIŃSKIE", "GE": "GRUZIŃSKIE",
    "MOŁDAWIA": "MOŁDAWSKIE", "MOLDOVA": "MOŁDAWSKIE", "MD": "MOŁDAWSKIE",
    "TURCJA": "TURECKIE", "TURKEY": "TURECKIE", "TÜRKIYE": "TURECKIE", "TR": "TURECKIE",
    "UZBEKISTAN": "UZBECKIE", "UZ": "UZBECKIE",
    "TADŻYKISTAN": "TADŻYCKIE", "TAJIKISTAN": "TADŻYCKIE", "TJ": "TADŻYCKIE",
    "KAZACHSTAN": "KAZACHSTAŃSKIE", "KAZAKHSTAN": "KAZACHSTAŃSKIE", "KZ": "KAZACHSTAŃSKIE",
    "KIRGISTAN": "KIRGISKIE", "KYRGYZSTAN": "KIRGISKIE", "KG": "KIRGISKIE",
    "TURKMENISTAN": "TURKMEŃSKIE", "TM": "TURKMEŃSKIE",
    "ARMENIA": "ARMEŃSKIE", "AM": "ARMEŃSKIE",
    "AZERBEJDŻAN": "AZERBEJDŻAŃSKIE", "AZERBAIJAN": "AZERBEJDŻAŃSKIE", "AZ": "AZERBEJDŻAŃSKIE",
    "PAKISTAN": "PAKISTAŃSKIE", "PK": "PAKISTAŃSKIE",
    "SRI LANKA": "LANKIJSKIE", "LK": "LANKIJSKIE",
    "CHINY": "CHIŃSKIE", "CHINA": "CHIŃSKIE", "CN": "CHIŃSKIE",
    "TAJLANDIA": "TAJSKIE", "THAILAND": "TAJSKIE", "TH": "TAJSKIE",
    "MALEZJA": "MALEZYJSKIE", "MALAYSIA": "MALEZYJSKIE", "MY": "MALEZYJSKIE",
    "ROSJA": "ROSYJSKIE", "RUSSIA": "ROSYJSKIE", "RU": "ROSYJSKIE",
    "NIEMCY": "NIEMIECKIE", "GERMANY": "NIEMIECKIE", "DE": "NIEMIECKIE",
    "FRANCJA": "FRANCUSKIE", "FRANCE": "FRANCUSKIE", "FR": "FRANCUSKIE",
    "WIELKA BRYTANIA": "BRYTYJSKIE", "UNITED KINGDOM": "BRYTYJSKIE", "GB": "BRYTYJSKIE",
    "WŁOCHY": "WŁOSKIE", "ITALY": "WŁOSKIE", "IT": "WŁOSKIE",
    "HISZPANIA": "HISZPAŃSKIE", "SPAIN": "HISZPAŃSKIE", "ES": "HISZPAŃSKIE",
    "PORTUGALIA": "PORTUGALSKIE", "PORTUGAL": "PORTUGALSKIE", "PT": "PORTUGALSKIE",
    "HOLANDIA": "HOLENDERSKIE", "NETHERLANDS": "HOLENDERSKIE", "NL": "HOLENDERSKIE",
    "BELGIA": "BELGIJSKIE", "BELGIUM": "BELGIJSKIE", "BE": "BELGIJSKIE",
    "CZECHY": "CZESKIE", "CZECH REPUBLIC": "CZESKIE", "CZECHIA": "CZESKIE", "CZ": "CZESKIE",
    "SŁOWACJA": "SŁOWACKIE", "SLOVAKIA": "SŁOWACKIE", "SK": "SŁOWACKIE",
    "LITWA": "LITEWSKIE", "LITHUANIA": "LITEWSKIE", "LT": "LITEWSKIE",
    "ŁOTWA": "ŁOTEWSKIE", "LATVIA": "ŁOTEWSKIE", "LV": "ŁOTEWSKIE",
    "ESTONIA": "ESTOŃSKIE", "EE": "ESTOŃSKIE",
    "RUMUNIA": "RUMUŃSKIE", "ROMANIA": "RUMUŃSKIE", "RO": "RUMUŃSKIE",
    "BUŁGARIA": "BUŁGARSKIE", "BULGARIA": "BUŁGARSKIE", "BG": "BUŁGARSKIE",
    "WĘGRY": "WĘGIERSKIE", "HUNGARY": "WĘGIERSKIE", "HU": "WĘGIERSKIE",
    "CHORWACJA": "CHORWACKIE", "CROATIA": "CHORWACKIE", "HR": "CHORWACKIE",
    "SERBIA": "SERBSKIE", "RS": "SERBSKIE",
    "BOŚNIA I HERCEGOWINA": "BOŚNIACKIE", "BOSNIA": "BOŚNIACKIE", "BA": "BOŚNIACKIE",
    "ALBANIA": "ALBAŃSKIE", "AL": "ALBAŃSKIE",
    "MACEDONIA PÓŁNOCNA": "MACEDOŃSKIE", "NORTH MACEDONIA": "MACEDOŃSKIE", "MK": "MACEDOŃSKIE",
    "KOSOWO": "KOSOWSKIE", "KOSOVO": "KOSOWSKIE", "XK": "KOSOWSKIE",
    "CZARNOGÓRA": "CZARNOGÓRSKIE", "MONTENEGRO": "CZARNOGÓRSKIE", "ME": "CZARNOGÓRSKIE",
    "SZWECJA": "SZWEDZKIE", "SWEDEN": "SZWEDZKIE", "SE": "SZWEDZKIE",
    "NORWEGIA": "NORWESKIE", "NORWAY": "NORWESKIE", "NO": "NORWESKIE",
    "DANIA": "DUŃSKIE", "DENMARK": "DUŃSKIE", "DK": "DUŃSKIE",
    "FINLANDIA": "FIŃSKIE", "FINLAND": "FIŃSKIE", "FI": "FIŃSKIE",
    "AUSTRIA": "AUSTRIACKIE", "AT": "AUSTRIACKIE",
    "SZWAJCARIA": "SZWAJCARSKIE", "SWITZERLAND": "SZWAJCARSKIE", "CH": "SZWAJCARSKIE",
    "IRLANDIA": "IRLANDZKIE", "IRELAND": "IRLANDZKIE", "IE": "IRLANDZKIE",
    "GRECJA": "GRECKIE", "GREECE": "GRECKIE", "GR": "GRECKIE",
    "SŁOWENIA": "SŁOWEŃSKIE", "SLOVENIA": "SŁOWEŃSKIE", "SI": "SŁOWEŃSKIE",
    "MEKSYK": "MEKSYKAŃSKIE", "MEXICO": "MEKSYKAŃSKIE", "MX": "MEKSYKAŃSKIE",
    "BRAZYLIA": "BRAZYLIJSKIE", "BRAZIL": "BRAZYLIJSKIE", "BR": "BRAZYLIJSKIE",
    "ARGENTYNA": "ARGENTYŃSKIE", "ARGENTINA": "ARGENTYŃSKIE", "AR": "ARGENTYŃSKIE",
    "PERU": "PERUWIAŃSKIE", "PE": "PERUWIAŃSKIE",
    "WENEZUELA": "WENEZUELSKIE", "VENEZUELA": "WENEZUELSKIE", "VE": "WENEZUELSKIE",
    "EKWADOR": "EKWADORSKIE", "ECUADOR": "EKWADORSKIE", "EC": "EKWADORSKIE",
    "BOLIWIA": "BOLIWIJSKIE", "BOLIVIA": "BOLIWIJSKIE", "BO": "BOLIWIJSKIE",
    "CHILE": "CHILIJSKIE", "CL": "CHILIJSKIE",
    "PARAGWAJ": "PARAGWAJSKIE", "PARAGUAY": "PARAGWAJSKIE", "PY": "PARAGWAJSKIE",
    "URUGWAJ": "URUGWAJSKIE", "URUGUAY": "URUGWAJSKIE", "UY": "URUGWAJSKIE",
    "KUBA": "KUBAŃSKIE", "CUBA": "KUBAŃSKIE", "CU": "KUBAŃSKIE",
    "NIGERIA": "NIGERYJSKIE", "NG": "NIGERYJSKIE",
    "ETIOPIA": "ETIOPSKIE", "ETHIOPIA": "ETIOPSKIE", "ET": "ETIOPSKIE",
    "EGIPT": "EGIPSKIE", "EGYPT": "EGIPSKIE", "EG": "EGIPSKIE",
    "MAROKO": "MAROKAŃSKIE", "MOROCCO": "MAROKAŃSKIE", "MA": "MAROKAŃSKIE",
    "TUNEZJA": "TUNEZYJSKIE", "TUNISIA": "TUNEZYJSKIE", "TN": "TUNEZYJSKIE",
    "ALGIERIA": "ALGIERSKIE", "ALGERIA": "ALGIERSKIE", "DZ": "ALGIERSKIE",
    "KOREA POŁUDNIOWA": "KOREAŃSKIE", "SOUTH KOREA": "KOREAŃSKIE", "KR": "KOREAŃSKIE",
    "JAPONIA": "JAPOŃSKIE", "JAPAN": "JAPOŃSKIE", "JP": "JAPOŃSKIE",
    "MONGOLIA": "MONGOLSKIE", "MN": "MONGOLSKIE",
    "MYANMAR": "MYANMARSKIE", "BIRMA": "BIRMAŃSKIE", "MM": "MYANMARSKIE",
    "KAMBODŻA": "KAMBODŻAŃSKIE", "CAMBODIA": "KAMBODŻAŃSKIE", "KH": "KAMBODŻAŃSKIE",
    "LAOS": "LAOTAŃSKIE", "LA": "LAOTAŃSKIE",
    "USA": "AMERYKAŃSKIE", "STANY ZJEDNOCZONE": "AMERYKAŃSKIE", "UNITED STATES": "AMERYKAŃSKIE", "US": "AMERYKAŃSKIE",
    "KANADA": "KANADYJSKIE", "CANADA": "KANADYJSKIE", "CA": "KANADYJSKIE",
    "AUSTRALIA": "AUSTRALIJSKIE", "AU": "AUSTRALIJSKIE",
    "NOWA ZELANDIA": "NOWOZELANDZKIE", "NEW ZEALAND": "NOWOZELANDZKIE", "NZ": "NOWOZELANDZKIE",
    "IZRAEL": "IZRAELSKIE", "ISRAEL": "IZRAELSKIE", "IL": "IZRAELSKIE",
    "IRAN": "IRAŃSKIE", "IR": "IRAŃSKIE",
    "IRAK": "IRACKIE", "IRAQ": "IRACKIE", "IQ": "IRACKIE",
    "SYRIA": "SYRYJSKIE", "SY": "SYRYJSKIE",
    "LIBAN": "LIBAŃSKIE", "LEBANON": "LIBAŃSKIE", "LB": "LIBAŃSKIE",
    "JORDANIA": "JORDAŃSKIE", "JORDAN": "JORDAŃSKIE", "JO": "JORDAŃSKIE",
    "ARABIA SAUDYJSKA": "SAUDYJSKIE", "SAUDI ARABIA": "SAUDYJSKIE", "SA": "SAUDYJSKIE",
}


def _odmien_obywatelstwo(kraj_lub_kod: str) -> str:
    """Konwertuje nazwę kraju lub kod na odmienione obywatelstwo"""
    if not kraj_lub_kod:
        return "POLSKIE"
    raw = str(kraj_lub_kod).strip().upper()
    # Jeśli już odmienione (kończy się na -SKIE, -CKIE, -ŃSKIE)
    if raw.endswith("SKIE") or raw.endswith("CKIE"):
        return raw
    # Szukaj w mapie
    result = _OBYWATELSTWO_MAP.get(raw)
    if result:
        return result
    # Fallback: zwróć jak jest (wielkie litery)
    logger.warning(f"Nieznane obywatelstwo: '{raw}' — wpisz odmienione ręcznie")
    return raw
def _fmt_data(data) -> str:
    if not data:
        return ""
    d = str(data).strip()
    # Obetnij timestamp "2026-06-10 00:00:00" → "2026-06-10"
    if " " in d:
        d = d.split(" ")[0]
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(d, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return d


# ============================================================
#  ROOT + ZAPIS
# ============================================================
def _utworz_root():
    ET.register_namespace("", KEDU_NS)
    ET.register_namespace("xsi", XSI_NS)
    root = ET.Element("KEDU")
    root.set("wersja_schematu", "1")
    root.set("xmlns:xsi", XSI_NS)
    root.set("xsi:schemaLocation", f"{KEDU_NS} {KEDU_XSD}")
    return root


def _zapisz_kedu(root, sciezka):
    sciezka = Path(sciezka)
    sciezka.parent.mkdir(parents=True, exist_ok=True)
    raw = ET.tostring(root, encoding="unicode", xml_declaration=False)
    raw = raw.replace(f'{{{KEDU_NS}}}', '')
    raw = raw.replace('<KEDU ', f'<KEDU xmlns="{KEDU_NS}" ', 1)
    full = f'<?xml version="1.0" encoding="UTF-8" ?>\n{raw}\n'
    sciezka.write_text(full, encoding="utf-8")
    logger.info(f"Zapisano KEDU: {sciezka}")
    return sciezka


# ============================================================
#  ZUSZUA — Zgłoszenie do ubezpieczeń
# ============================================================
def _buduj_zuszua(root, pracownicy, doc_id_start=1):
    """
    ZUSZUA — jeden dokument per pracownik.
    Struktura wg Enova: I-XIV.
    """
    poprawni = [p for p in pracownicy if not p.bledy]

    for idx, p in enumerate(poprawni):
        zua = ET.SubElement(root, "ZUSZUA")
        zua.set("id_dokumentu", str(doc_id_start + idx))

        # I — Typ zgłoszenia
        sek_i = _el(zua, "I")
        _el(sek_i, "p1", "true")

        # II — Dane płatnika
        _buduj_platnik(zua)

        # III — Dane ubezpieczonego
        sek_iii = _el(zua, "III")
        if p.pesel:
            _el(sek_iii, "p1", str(p.pesel))
        _el(sek_iii, "p5", p.nazwisko.upper())
        _el(sek_iii, "p6", p.imie.upper())
        if p.data_urodzenia:
            _el(sek_iii, "p7", _fmt_data(p.data_urodzenia))

        # IV — Obywatelstwo i płeć
        sek_iv = _el(zua, "IV")
        obyw = _odmien_obywatelstwo(p.obywatelstwo)
        _el(sek_iv, "p3", obyw)
        plec = ""
        if p.pesel and len(p.pesel) >= 10:
            try:
                plec = "K" if int(p.pesel[9]) % 2 == 0 else "M"
            except ValueError:
                pass
        if plec:
            _el(sek_iv, "p4", plec)

        # V — Kod tytułu ubezpieczenia
        sek_v = _el(zua, "V")
        p1v = _el(sek_v, "p1")
        _el(p1v, "p1", str(p.kod_tytulu).zfill(4))
        _el(p1v, "p2", "0")
        _el(p1v, "p3", "0")

        # VI — Data zgłoszenia + rodzaje ubezpieczeń
        sek_vi = _el(zua, "VI")
        data_zgl = _fmt_data(p.data_zgłoszenia) or datetime.now().strftime("%Y-%m-%d")
        _el(sek_vi, "p1", data_zgl)
        _el(sek_vi, "p2", "true" if p.ubezp_emerytalne else "false")
        _el(sek_vi, "p3", "true" if p.ubezp_rentowe else "false")
        _el(sek_vi, "p4", "true" if p.ubezp_chorobowe else "false")
        _el(sek_vi, "p5", "true" if p.ubezp_wypadkowe else "false")

        # VII — Ubezpieczenie zdrowotne
        sek_vii = _el(zua, "VII")
        _el(sek_vii, "p1", data_zgl)
        kod_nfz = str(p.kod_nfz).zfill(2) if p.kod_nfz else "07"
        _el(sek_vii, "p2", kod_nfz + "R")

        # VIII-IX — puste
        _el(zua, "VIII")
        _el(zua, "IX")

        # X — TERYT (opcjonalne)
        sek_x = _el(zua, "X")
        _el(sek_x, "p3")

        # XI — Adres
        sek_xi = _el(zua, "XI")
        if p.kod_pocztowy:
            _el(sek_xi, "p1", str(p.kod_pocztowy).replace("-", ""))
        if p.miejscowosc:
            _el(sek_xi, "p2", p.miejscowosc.upper())
            _el(sek_xi, "p3", p.miejscowosc.upper())
        if p.ulica:
            _el(sek_xi, "p4", p.ulica.upper())
        if p.nr_domu:
            _el(sek_xi, "p5", str(p.nr_domu))
        if p.nr_lokalu:
            _el(sek_xi, "p6", str(p.nr_lokalu))

        # XII-XIII — puste
        _el(zua, "XII")
        _el(zua, "XIII")

        # XIV — Data wypełnienia
        sek_xiv = _el(zua, "XIV")
        _el(sek_xiv, "p1", datetime.now().strftime("%Y-%m-%d"))

    return len(poprawni)


# ============================================================
#  ZUSZZA — Zgłoszenie tylko do ubezpieczenia zdrowotnego
# ============================================================
def _buduj_zuszza(root, pracownicy, doc_id_start=1):
    """
    ZUSZZA — zgłoszenie do ubezpieczenia zdrowotnego.
    Struktura wg Enova (witkowska_zza.xml):
      I    - typ (true)
      II   - dane płatnika
      III  - dane ubezpieczonego
      IV   - obywatelstwo, płeć
      V    - kod tytułu + TERYT
      VI   - ubezp. zdrowotne (data + NFZ)
      VII  - puste
      VIII - adres zamieszkania
      IX   - adres korespondencyjny
      X    - puste
      XI   - data wypełnienia
    """
    poprawni = [p for p in pracownicy if not p.bledy]

    for idx, p in enumerate(poprawni):
        zza = ET.SubElement(root, "ZUSZZA")
        zza.set("id_dokumentu", str(doc_id_start + idx))

        # I
        sek_i = _el(zza, "I")
        _el(sek_i, "p1", "true")

        # II — Dane płatnika
        _buduj_platnik(zza)

        # III — Dane ubezpieczonego
        sek_iii = _el(zza, "III")
        if p.pesel:
            _el(sek_iii, "p1", str(p.pesel))
        _el(sek_iii, "p5", p.nazwisko.upper())
        _el(sek_iii, "p6", p.imie.upper())
        if p.data_urodzenia:
            _el(sek_iii, "p7", _fmt_data(p.data_urodzenia))

        # IV — Obywatelstwo i płeć
        sek_iv = _el(zza, "IV")
        _el(sek_iv, "p3", _odmien_obywatelstwo(p.obywatelstwo))
        plec = ""
        if p.pesel and len(p.pesel) >= 10:
            try:
                plec = "K" if int(p.pesel[9]) % 2 == 0 else "M"
            except ValueError:
                pass
        if plec:
            _el(sek_iv, "p4", plec)

        # V — Kod tytułu
        sek_v = _el(zza, "V")
        p1v = _el(sek_v, "p1")
        _el(p1v, "p1", str(p.kod_tytulu).zfill(4))
        _el(p1v, "p2", "0")
        _el(p1v, "p3", "0")

        # VI — Ubezp. zdrowotne
        sek_vi = _el(zza, "VI")
        data_zgl = _fmt_data(p.data_zgłoszenia) or datetime.now().strftime("%Y-%m-%d")
        _el(sek_vi, "p1", data_zgl)
        kod_nfz = str(p.kod_nfz).zfill(2) if p.kod_nfz else "07"
        _el(sek_vi, "p2", kod_nfz + "R")

        # VII — puste
        _el(zza, "VII")

        # VIII — Adres zamieszkania
        sek_viii = _el(zza, "VIII")
        if p.kod_pocztowy:
            _el(sek_viii, "p1", str(p.kod_pocztowy).replace("-", ""))
        if p.miejscowosc:
            _el(sek_viii, "p2", p.miejscowosc.upper())
            _el(sek_viii, "p3", p.miejscowosc.upper())
        if p.ulica:
            _el(sek_viii, "p4", p.ulica.upper())
        if p.nr_domu:
            _el(sek_viii, "p5", str(p.nr_domu))
        if p.nr_lokalu:
            _el(sek_viii, "p6", str(p.nr_lokalu))

        # IX — Adres korespondencyjny (taki sam jak zamieszkania)
        sek_ix = _el(zza, "IX")
        if p.kod_pocztowy:
            _el(sek_ix, "p1", str(p.kod_pocztowy).replace("-", ""))
        if p.miejscowosc:
            _el(sek_ix, "p2", p.miejscowosc.upper())
        if p.ulica:
            _el(sek_ix, "p4", p.ulica.upper())
        if p.nr_domu:
            _el(sek_ix, "p5", str(p.nr_domu))
        if p.nr_lokalu:
            _el(sek_ix, "p6", str(p.nr_lokalu))

        # X — puste
        _el(zza, "X")

        # XI — Data wypełnienia
        sek_xi = _el(zza, "XI")
        _el(sek_xi, "p1", datetime.now().strftime("%Y-%m-%d"))

    return len(poprawni)


# ============================================================
#  ZWUA — Wyrejestrowanie z ubezpieczeń
# ============================================================
def _buduj_zwua(root, pracownicy, doc_id_start=1):
    """
    ZWUA — wyrejestrowanie z ubezpieczeń.
    Struktura wg Enova (wiejak_ZWUA.xml):
      I   - typ (true)
      II  - dane płatnika
      III - dane ubezpieczonego
      IV  - wyrejestrowanie z ubezp. społecznych (kod tytułu + data + przyczyna)
      V   - wyrejestrowanie z ubezp. zdrowotnego (data + NFZ + kod przyczyny)
      VI  - data wypełnienia
    """
    poprawni = [p for p in pracownicy if not p.bledy]

    for idx, p in enumerate(poprawni):
        zwua = ET.SubElement(root, "ZUSZWUA")
        zwua.set("id_dokumentu", str(doc_id_start + idx))

        # I
        sek_i = _el(zwua, "I")
        _el(sek_i, "p1", "true")

        # II — Dane płatnika
        _buduj_platnik(zwua)

        # III — Dane ubezpieczonego
        sek_iii = _el(zwua, "III")
        if p.pesel:
            _el(sek_iii, "p1", str(p.pesel))
        _el(sek_iii, "p5", p.nazwisko.upper())
        _el(sek_iii, "p6", p.imie.upper())
        if p.data_urodzenia:
            _el(sek_iii, "p7", _fmt_data(p.data_urodzenia))

        # IV — Wyrejestrowanie z ubezpieczeń społecznych
        sek_iv = _el(zwua, "IV")
        p1v = _el(sek_iv, "p1")
        _el(p1v, "p1", str(p.kod_tytulu).zfill(4))
        _el(p1v, "p2", "0")
        _el(p1v, "p3", "0")
        data_wyr = _fmt_data(p.data_zgłoszenia) or datetime.now().strftime("%Y-%m-%d")
        _el(sek_iv, "p2", data_wyr)
        kod_przyczyny = str(getattr(p, 'kod_wyrejestrowania', '100')).strip() or "100"
        _el(sek_iv, "p3", kod_przyczyny)

        # V — Wyrejestrowanie z ubezp. zdrowotnego
        sek_v = _el(zwua, "V")
        _el(sek_v, "p1", data_wyr)
        kod_nfz = str(p.kod_nfz).zfill(2) if p.kod_nfz else "07"
        _el(sek_v, "p2", kod_nfz + "R")
        _el(sek_v, "p3", "402")
        _el(sek_v, "p5", "1")

        # VI — Data wypełnienia
        sek_vi = _el(zwua, "VI")
        _el(sek_vi, "p1", datetime.now().strftime("%Y-%m-%d"))

    return len(poprawni)


# ============================================================
#  ZIUA — Zmiana danych identyfikacyjnych ubezpieczonego
# ============================================================
def _buduj_ziua(root, pracownicy, doc_id_start=1):
    """ZIUA — zmiana danych identyfikacyjnych (np. paszport → PESEL)."""
    poprawni = [p for p in pracownicy if not p.bledy]

    for idx, p in enumerate(poprawni):
        ziua = ET.SubElement(root, "ZUSZIUA")
        ziua.set("id_dokumentu", str(doc_id_start + idx))

        # I — Dane płatnika
        _buduj_platnik(ziua)

        # II — Poprzednie dane identyfikacyjne
        sek_ii = _el(ziua, "II")
        # Poprzedni identyfikator
        if p.poprzedni_typ_id == "2" and p.poprzedni_nr_dokumentu:
            # Był paszport
            _el(sek_ii, "p4", "2")  # typ: paszport
            _el(sek_ii, "p5", str(p.poprzedni_nr_dokumentu))
        elif p.poprzedni_nr_dokumentu:
            _el(sek_ii, "p4", str(p.poprzedni_typ_id or "1"))
            _el(sek_ii, "p5", str(p.poprzedni_nr_dokumentu))
        _el(sek_ii, "p6", p.nazwisko.upper())
        _el(sek_ii, "p7", p.imie.upper())
        if p.data_urodzenia:
            _el(sek_ii, "p8", _fmt_data(p.data_urodzenia))

        # III — Aktualne (nowe) dane identyfikacyjne
        sek_iii = _el(ziua, "III")
        if p.pesel:
            _el(sek_iii, "p1", str(p.pesel))
        _el(sek_iii, "p6", p.nazwisko.upper())
        _el(sek_iii, "p7", p.imie.upper())
        if p.data_urodzenia:
            _el(sek_iii, "p8", _fmt_data(p.data_urodzenia))

        # IV — Data wypełnienia
        sek_iv = _el(ziua, "IV")
        _el(sek_iv, "p1", datetime.now().strftime("%Y-%m-%d"))

    return len(poprawni)


# ============================================================
#  GŁÓWNA FUNKCJA
# ============================================================
def generuj_wszystkie_kedu(pracownicy, katalog):
    katalog = Path(katalog)
    katalog.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    wyniki = {}

    rca = [p for p in pracownicy if p.typ_deklaracji.upper() == "RCA" and not p.bledy]
    zua = [p for p in pracownicy if p.typ_deklaracji.upper() == "ZUA" and not p.bledy]
    zza = [p for p in pracownicy if p.typ_deklaracji.upper() == "ZZA" and not p.bledy]
    zwua = [p for p in pracownicy if p.typ_deklaracji.upper() == "ZWUA" and not p.bledy]
    ziua = [p for p in pracownicy if p.typ_deklaracji.upper() == "ZIUA" and not p.bledy]

    if rca:
        root = _utworz_root()
        _buduj_naglowek(root)
        _buduj_zusdra(root, rca, doc_id=1)
        _buduj_zusrca(root, rca, doc_id=2)
        sciezka = katalog / f"DRA_RCA_{ts}.xml"
        _zapisz_kedu(root, sciezka)
        wyniki["DRA+RCA"] = sciezka
        logger.info(f"DRA+RCA: {len(rca)} pracowników")

    if zua:
        root = _utworz_root()
        _buduj_naglowek(root)
        count = _buduj_zuszua(root, zua, doc_id_start=1)
        sciezka = katalog / f"ZUA_{ts}.xml"
        _zapisz_kedu(root, sciezka)
        wyniki["ZUA"] = sciezka
        logger.info(f"ZUA: {count} zgłoszeń")

    if zza:
        root = _utworz_root()
        _buduj_naglowek(root)
        count = _buduj_zuszza(root, zza, doc_id_start=1)
        sciezka = katalog / f"ZZA_{ts}.xml"
        _zapisz_kedu(root, sciezka)
        wyniki["ZZA"] = sciezka
        logger.info(f"ZZA: {count} zgłoszeń")

    if zwua:
        root = _utworz_root()
        _buduj_naglowek(root)
        count = _buduj_zwua(root, zwua, doc_id_start=1)
        sciezka = katalog / f"ZWUA_{ts}.xml"
        _zapisz_kedu(root, sciezka)
        wyniki["ZWUA"] = sciezka
        logger.info(f"ZWUA: {count} wyrejestrowań")

    if ziua:
        root = _utworz_root()
        _buduj_naglowek(root)
        count = _buduj_ziua(root, ziua, doc_id_start=1)
        sciezka = katalog / f"ZIUA_{ts}.xml"
        _zapisz_kedu(root, sciezka)
        wyniki["ZIUA"] = sciezka
        logger.info(f"ZIUA: {count} zmian danych")

    return wyniki
    """
    ZUSZZA — zgłoszenie TYLKO do ubezpieczenia zdrowotnego.
    Uproszczona wersja ZUA — bez składek społecznych.
    """
    poprawni = [p for p in pracownicy if not p.bledy]

    for idx, p in enumerate(poprawni):
        zza = ET.SubElement(root, "ZUSZZA")
        zza.set("id_dokumentu", str(doc_id_start + idx))

        # I — Typ zgłoszenia
        sek_i = _el(zza, "I")
        _el(sek_i, "p1", "true")

        # II — Dane płatnika
        _buduj_platnik(zza)

        # III — Dane ubezpieczonego
        sek_iii = _el(zza, "III")
        if p.pesel:
            _el(sek_iii, "p1", str(p.pesel))
        _el(sek_iii, "p5", p.nazwisko.upper())
        _el(sek_iii, "p6", p.imie.upper())
        if p.data_urodzenia:
            _el(sek_iii, "p7", _fmt_data(p.data_urodzenia))

        # IV — Obywatelstwo i płeć
        sek_iv = _el(zza, "IV")
        obyw = _odmien_obywatelstwo(p.obywatelstwo)
        _el(sek_iv, "p3", obyw)
        plec = ""
        if p.pesel and len(p.pesel) >= 10:
            try:
                plec = "K" if int(p.pesel[9]) % 2 == 0 else "M"
            except ValueError:
                pass
        if plec:
            _el(sek_iv, "p4", plec)

        # V — Kod tytułu ubezpieczenia
        sek_v = _el(zza, "V")
        p1v = _el(sek_v, "p1")
        _el(p1v, "p1", str(p.kod_tytulu).zfill(4))
        _el(p1v, "p2", "0")
        _el(p1v, "p3", "0")

        # VI — Data zgłoszenia do ubezp. zdrowotnego + kod NFZ
        sek_vi = _el(zza, "VI")
        data_zgl = _fmt_data(p.data_zgłoszenia) or datetime.now().strftime("%Y-%m-%d")
        _el(sek_vi, "p1", data_zgl)
        kod_nfz = str(p.kod_nfz).zfill(2) if p.kod_nfz else "07"
        _el(sek_vi, "p2", kod_nfz + "R")

        # VII-IX — puste
        _el(zza, "VII")
        _el(zza, "VIII")
        _el(zza, "IX")

        # X — TERYT
        sek_x = _el(zza, "X")
        _el(sek_x, "p3")

        # XI — Adres
        sek_xi = _el(zza, "XI")
        if p.kod_pocztowy:
            _el(sek_xi, "p1", str(p.kod_pocztowy).replace("-", ""))
        if p.miejscowosc:
            _el(sek_xi, "p2", p.miejscowosc.upper())
            _el(sek_xi, "p3", p.miejscowosc.upper())
        if p.ulica:
            _el(sek_xi, "p4", p.ulica.upper())
        if p.nr_domu:
            _el(sek_xi, "p5", str(p.nr_domu))
        if p.nr_lokalu:
            _el(sek_xi, "p6", str(p.nr_lokalu))

        # XII-XIII — puste
        _el(zza, "XII")
        _el(zza, "XIII")

        # XIV — Data wypełnienia
        sek_xiv = _el(zza, "XIV")
        _el(sek_xiv, "p1", datetime.now().strftime("%Y-%m-%d"))

    return len(poprawni)


def generuj_wszystkie_kedu(pracownicy, katalog):
    katalog = Path(katalog)
    katalog.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    wyniki = {}

    rca = [p for p in pracownicy if p.typ_deklaracji.upper() == "RCA" and not p.bledy]
    zua = [p for p in pracownicy if p.typ_deklaracji.upper() == "ZUA" and not p.bledy]
    zza = [p for p in pracownicy if p.typ_deklaracji.upper() == "ZZA" and not p.bledy]

    if rca:
        root = _utworz_root()
        _buduj_naglowek(root)
        _buduj_zusdra(root, rca, doc_id=1)
        _buduj_zusrca(root, rca, doc_id=2)
        sciezka = katalog / f"DRA_RCA_{ts}.xml"
        _zapisz_kedu(root, sciezka)
        wyniki["DRA+RCA"] = sciezka
        logger.info(f"DRA+RCA: {len(rca)} pracowników")

    if zua:
        root = _utworz_root()
        _buduj_naglowek(root)
        count = _buduj_zuszua(root, zua, doc_id_start=1)
        sciezka = katalog / f"ZUA_{ts}.xml"
        _zapisz_kedu(root, sciezka)
        wyniki["ZUA"] = sciezka
        logger.info(f"ZUA: {count} zgłoszeń")

    if zza:
        root = _utworz_root()
        _buduj_naglowek(root)
        count = _buduj_zuszza(root, zza, doc_id_start=1)
        sciezka = katalog / f"ZZA_{ts}.xml"
        _zapisz_kedu(root, sciezka)
        wyniki["ZZA"] = sciezka
        logger.info(f"ZZA: {count} zgłoszeń")

    return wyniki
