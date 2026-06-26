"""
Skaner dokumentów — odczyt danych z paszportów i powiadomień o nadaniu PESEL.
Używa Claude Vision API do ekstrakcji danych.

Obsługuje:
  - Paszporty (zdjęcie/skan) → imię, nazwisko, nr paszportu, data ur., obywatelstwo, płeć
  - Powiadomienie o nadaniu PESEL → PESEL, imię, nazwisko, data ur.
  - Dowody osobiste → dane identyfikacyjne
"""
import json
import base64
import logging
from pathlib import Path
from dataclasses import dataclass

logger = logging.getLogger("platnik_auto")

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False


@dataclass
class DaneZDokumentu:
    """Dane wyekstrahowane z dokumentu"""
    typ_dokumentu: str = ""      # "paszport", "pesel", "dowod"
    imie: str = ""
    nazwisko: str = ""
    drugie_imie: str = ""
    pesel: str = ""
    nr_dokumentu: str = ""       # nr paszportu / dowodu
    seria_dokumentu: str = ""
    kraj_wydania: str = ""       # kod kraju np. "PHL", "COL"
    obywatelstwo: str = ""       # np. "Filipiny", "Kolumbia"
    data_urodzenia: str = ""     # YYYY-MM-DD
    plec: str = ""               # "K" lub "M"
    data_waznosci: str = ""      # data ważności dokumentu
    mrz: str = ""                # surowy tekst MRZ (jeśli odczytany)
    pewnosc: str = ""            # "wysoka", "średnia", "niska"
    uwagi: str = ""              # dodatkowe uwagi z OCR


PROMPT_PASZPORT = """Analizujesz zdjęcie/skan paszportu. Wyodrębnij WSZYSTKIE dane i zwróć TYLKO JSON (bez markdown, bez komentarzy):

{
  "typ_dokumentu": "paszport",
  "imie": "IMIĘ (pierwsze, łacińskie litery)",
  "nazwisko": "NAZWISKO (łacińskie litery)",
  "drugie_imie": "DRUGIE IMIĘ (jeśli jest)",
  "nr_dokumentu": "NUMER PASZPORTU",
  "kraj_wydania": "KRAJ WYDANIA (pełna nazwa po polsku, np. Filipiny, Kolumbia)",
  "obywatelstwo": "OBYWATELSTWO (pełna nazwa po polsku)",
  "data_urodzenia": "DATA URODZENIA w formacie YYYY-MM-DD",
  "plec": "K lub M",
  "data_waznosci": "DATA WAŻNOŚCI w formacie YYYY-MM-DD",
  "mrz": "TEKST MRZ (dwie dolne linie z paszportu, jeśli widoczne)",
  "pewnosc": "wysoka/średnia/niska — jak pewny jesteś odczytu",
  "uwagi": "dodatkowe uwagi, np. nieczytelne fragmenty"
}

WAŻNE:
- Imiona i nazwiska zapisz WIELKIMI LITERAMI bez znaków diakrytycznych specyficznych dla danego języka (ñ→N, é→E, á→A). Zachowaj polskie znaki jeśli występują.
- Kraj wydania i obywatelstwo po polsku (np. Filipiny, nie Philippines).
- Jeśli nie możesz odczytać pola — wpisz pusty string "".
- Zwróć TYLKO JSON, nic więcej."""

PROMPT_PESEL = """Analizujesz zdjęcie/skan powiadomienia o nadaniu numeru PESEL. Wyodrębnij dane i zwróć TYLKO JSON:

{
  "typ_dokumentu": "pesel",
  "pesel": "NUMER PESEL (11 cyfr)",
  "imie": "IMIĘ",
  "nazwisko": "NAZWISKO",
  "drugie_imie": "DRUGIE IMIĘ (jeśli jest)",
  "data_urodzenia": "DATA URODZENIA w formacie YYYY-MM-DD",
  "plec": "K lub M",
  "obywatelstwo": "OBYWATELSTWO (jeśli podane)",
  "pewnosc": "wysoka/średnia/niska",
  "uwagi": "dodatkowe uwagi"
}

WAŻNE: Zwróć TYLKO JSON, nic więcej. Imiona WIELKIMI LITERAMI."""

PROMPT_AUTO = """Analizujesz zdjęcie dokumentu tożsamości. Najpierw rozpoznaj typ dokumentu (paszport, dowód osobisty, powiadomienie o nadaniu PESEL, inny dokument), a potem wyodrębnij dane.

Zwróć TYLKO JSON:

{
  "typ_dokumentu": "paszport / pesel / dowod / inny",
  "imie": "IMIĘ",
  "nazwisko": "NAZWISKO",
  "drugie_imie": "",
  "pesel": "PESEL (jeśli widoczny)",
  "nr_dokumentu": "NUMER DOKUMENTU",
  "seria_dokumentu": "SERIA (jeśli dotyczy)",
  "kraj_wydania": "KRAJ po polsku",
  "obywatelstwo": "OBYWATELSTWO po polsku",
  "data_urodzenia": "YYYY-MM-DD",
  "plec": "K lub M",
  "data_waznosci": "YYYY-MM-DD (jeśli dotyczy)",
  "mrz": "tekst MRZ (jeśli widoczny)",
  "pewnosc": "wysoka/średnia/niska",
  "uwagi": ""
}

Imiona WIELKIMI LITERAMI, bez obcych znaków diakrytycznych (ñ→N, é→E). Zwróć TYLKO JSON."""


def skanuj_dokument(
    image_path: str | Path = None,
    image_bytes: bytes = None,
    image_base64: str = None,
    typ: str = "auto",
    api_key: str = "",
) -> DaneZDokumentu:
    """
    Skanuje dokument (paszport/PESEL) i zwraca wyekstrahowane dane.
    
    Args:
        image_path: ścieżka do pliku obrazu
        image_bytes: surowe bajty obrazu
        image_base64: obraz jako base64
        typ: "paszport", "pesel", "auto" (autodetect)
        api_key: klucz API Anthropic
    """
    if not REQUESTS_AVAILABLE:
        raise RuntimeError("Moduł requests nie jest zainstalowany")
    
    if not api_key:
        raise ValueError("Brak klucza API Anthropic. Wpisz go w ustawieniach.")

    # Przygotuj obraz jako base64
    if image_path:
        path = Path(image_path)
        image_bytes = path.read_bytes()
    
    if image_bytes:
        image_base64 = base64.standard_b64encode(image_bytes).decode("utf-8")
    
    if not image_base64:
        raise ValueError("Brak obrazu do analizy")

    # Wykryj MIME type
    if image_bytes:
        if image_bytes[:4] == b'\x89PNG':
            media_type = "image/png"
        elif image_bytes[:2] == b'\xff\xd8':
            media_type = "image/jpeg"
        elif image_bytes[:4] == b'%PDF':
            media_type = "application/pdf"
        elif image_bytes[:4] in (b'RIFF', b'WEBP'):
            media_type = "image/webp"
        else:
            media_type = "image/jpeg"
    else:
        media_type = "image/jpeg"

    # Wybierz prompt
    if typ == "paszport":
        prompt = PROMPT_PASZPORT
    elif typ == "pesel":
        prompt = PROMPT_PESEL
    else:
        prompt = PROMPT_AUTO

    # Wywołaj Claude Vision API
    logger.info(f"Skanuję dokument (typ={typ})...")
    
    payload = {
        "model": "claude-sonnet-4-6",
        "max_tokens": 1000,
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": media_type,
                            "data": image_base64,
                        }
                    },
                    {
                        "type": "text",
                        "text": prompt,
                    }
                ]
            }
        ]
    }

    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "Content-Type": "application/json",
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
            },
            json=payload,
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
    except requests.RequestException as e:
        raise RuntimeError(f"Błąd API: {e}")

    # Parsuj odpowiedź
    text = ""
    for block in data.get("content", []):
        if block.get("type") == "text":
            text += block["text"]

    # Wyciągnij JSON
    text = text.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    text = text.strip()

    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        logger.error(f"Nie udało się sparsować odpowiedzi: {text[:200]}")
        return DaneZDokumentu(uwagi=f"Błąd parsowania: {text[:200]}")

    # Mapuj na DaneZDokumentu
    result = DaneZDokumentu(
        typ_dokumentu=parsed.get("typ_dokumentu", ""),
        imie=parsed.get("imie", ""),
        nazwisko=parsed.get("nazwisko", ""),
        drugie_imie=parsed.get("drugie_imie", ""),
        pesel=parsed.get("pesel", ""),
        nr_dokumentu=parsed.get("nr_dokumentu", ""),
        seria_dokumentu=parsed.get("seria_dokumentu", ""),
        kraj_wydania=parsed.get("kraj_wydania", ""),
        obywatelstwo=parsed.get("obywatelstwo", ""),
        data_urodzenia=parsed.get("data_urodzenia", ""),
        plec=parsed.get("plec", ""),
        data_waznosci=parsed.get("data_waznosci", ""),
        mrz=parsed.get("mrz", ""),
        pewnosc=parsed.get("pewnosc", ""),
        uwagi=parsed.get("uwagi", ""),
    )

    logger.info(f"Odczytano: {result.imie} {result.nazwisko} ({result.typ_dokumentu}, pewność: {result.pewnosc})")
    return result


def skanuj_wiele(images: list[tuple[str, bytes]], typ: str = "auto", api_key: str = "") -> list[DaneZDokumentu]:
    """Skanuje wiele dokumentów naraz"""
    wyniki = []
    for nazwa, img_bytes in images:
        try:
            wynik = skanuj_dokument(image_bytes=img_bytes, typ=typ, api_key=api_key)
            wyniki.append(wynik)
            logger.info(f"  {nazwa}: {wynik.imie} {wynik.nazwisko}")
        except Exception as e:
            logger.error(f"  {nazwa}: błąd — {e}")
            wyniki.append(DaneZDokumentu(uwagi=f"Błąd: {e}"))
    return wyniki
