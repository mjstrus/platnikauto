"""
Płatnik Auto — Streamlit App v3
Upload Excel → edytuj dane → wybierz operację → pobierz KEDU + PDF

Uruchomienie:
  streamlit run app.py
"""
import streamlit as st
import pandas as pd
import tempfile
import zipfile
import io
import sys
from pathlib import Path
from datetime import datetime, date
import requests

sys.path.insert(0, str(Path(__file__).parent))

import config
from excel_reader import wczytaj_excel, raport_walidacji, kod_pocztowy_do_nfz, waliduj_pesel
from kedu_generator import generuj_wszystkie_kedu
from pdf_confirmations import generuj_potwierdzenia
from document_scanner import skanuj_dokument, DaneZDokumentu
from sheets_connector import SheetsConnector, GSPREAD_AVAILABLE



# === POBIERANIE DANYCH PO NIP ===
def pobierz_dane_z_nip(nip: str) -> dict | None:
    """Pobiera dane firmy z Białej Listy VAT (MF) lub REGON (GUS)"""
    nip = nip.strip().replace("-", "")
    if len(nip) != 10:
        return None
    
    # Próba 1: Biała Lista VAT (Ministerstwo Finansów) — darmowe, bez klucza
    try:
        today = date.today().strftime("%Y-%m-%d")
        url = f"https://wl-api.mf.gov.pl/api/search/nip/{nip}?date={today}"
        resp = requests.get(url, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            subject = data.get("result", {}).get("subject")
            if subject:
                nazwa = subject.get("name", "")
                regon = subject.get("regon", "")
                adres = subject.get("residenceAddress") or subject.get("workingAddress") or ""
                return {
                    "nazwa": nazwa,
                    "regon": regon,
                    "adres": adres,
                    "status_vat": subject.get("statusVat", ""),
                    "zrodlo": "Biała Lista VAT (MF)",
                }
    except Exception:
        pass
    
    # Próba 2: KRS API (Ministerstwo Sprawiedliwości) — darmowe
    try:
        url = f"https://api-krs.ms.gov.pl/api/krs/OdpisAktualny?nip={nip}&rejestr=P&format=json"
        resp = requests.get(url, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            odpis = data.get("odpis", {})
            dane = odpis.get("dane", {}).get("dzial1", {})
            nazwa = dane.get("danePodmiotu", {}).get("nazwa", "")
            adres_d = dane.get("siedzibaIAdres", {}).get("adres", {})
            regon = dane.get("danePodmiotu", {}).get("identyfikatory", {}).get("regon", "")
            adres = f"{adres_d.get('ulica', '')} {adres_d.get('nrDomu', '')}, {adres_d.get('kodPocztowy', '')} {adres_d.get('miejscowosc', '')}".strip(", ")
            if nazwa:
                return {
                    "nazwa": nazwa,
                    "regon": regon,
                    "adres": adres,
                    "status_vat": "",
                    "zrodlo": "KRS (MS)",
                }
    except Exception:
        pass
    
    return None

# === PAGE CONFIG ===
st.set_page_config(page_title="Płatnik Auto", page_icon="📋", layout="wide")

st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; }
    div[data-testid="stMetric"] { background: #0f172a; padding: 10px 14px; border-radius: 8px; border-left: 3px solid #3b82f6; }
</style>
""", unsafe_allow_html=True)

st.markdown("# 📋 Płatnik Auto")
st.markdown("*Excel → edycja → KEDU + potwierdzenia PDF*")
st.divider()

# === SESSION STATE ===
for key, default in [
    ("df_edit", None), ("pracownicy", []),
    ("kedu_bytes", None), ("pdf_zip_bytes", None),
    ("kedu_filename", ""), ("loaded", False),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ============================================================
#  SIDEBAR
# ============================================================
KODY_TYTULU_REF = {
    "01xx": "PRACOWNICY: 0110-pracownik, 0111-młodociany, 0112-świadcz.szkol., 0120-służba zastępcza, 0125-praktykant, 0126-tymczasowy",
    "04xx": "ZLECENIOBIORCY: 0411-zlecenie pełne, 0417-tylko zdrowotne, 0426-niania, 0428-rada nadzorcza, 0429-zlecenie zagr., 0431-niania nadwyżka",
    "05xx": "DZIAŁALNOŚĆ: 0510-pełne, 0511-współpracujący, 0512-z rentą, 0514-wakacje, 0540-preferencyjne, 0570-ulga na start, 0580-Mały ZUS, 0590-Mały ZUS Plus",
    "06-09": "INNE: 0610-poseł, 0700-stypendysta, 0800-duchowny, 0860-pomocnik rolnika, 0920-staż PUP, 0921-niepełnospr. PUP",
    "11-19": "SZCZEGÓLNE: 1110-żołnierz, 1211-urlop wych., 1240-zasiłek macierz., 1813-student, 1910-rolnik KRUS",
    "20-25": "ŚWIADCZENIA: 2000-emeryt/rencista, 2110-św.pielęgnacyjne, 2241-rada nadzorcza, 2250-powołany/prokurent, 2400-marynarz, 2500-niania budżet",
    "30xx": "ZDROWOTNE: 3000-inna osoba objęta ubezp. zdrowotnym",
}

with st.sidebar:
    # --- OPERACJA ---
    st.header("🎯 Rodzaj operacji")

    operacja = st.selectbox("Co chcesz zrobić?", [
        "ZUA — Zgłoszenie do ubezpieczeń",
        "ZZA — Zgłoszenie tylko zdrowotne",
        "ZWUA — Wyrejestrowanie z ubezpieczeń",
        "RCA — Raport rozliczeniowy miesięczny",
        "ZIUA — Zmiana danych identyfikacyjnych",
    ], help="Wybierz typ deklaracji do wygenerowania")

    typ_operacji = operacja.split(" — ")[0]  # ZUA, ZZA, ZWUA, RCA, ZIUA

    st.divider()

    # --- PARAMETRY OPERACJI ---
    st.header("📋 Parametry")

    if typ_operacji in ("ZUA", "ZZA"):
        data_zgloszenia = st.date_input("Data zgłoszenia", value=date.today())
        kod_tytulu_input = st.text_input("Kod tytułu ubezpieczenia (4 cyfry)", value="0110", max_chars=4, help="Wpisz 4-cyfrowy kod, np. 0110, 0411, 2250")
        with st.expander("📖 Podpowiedź — popularne kody"):
            for grupa, opis in KODY_TYTULU_REF.items():
                st.caption(f"**{grupa}**: {opis}")

        KODY_ZAWODOW = [
            "932101 — Pakowacz ręczny",
            "932990 — Pracownik produkcji (prace proste)",
            "932901 — Robotnik gospodarczy",
            "411003 — Pracownik biurowy",
            "242490 — Specjalista ds. rekrutacji / koordynator",
            "833101 — Kierowca samochodu ciężarowego",
            "833203 — Kierowca samochodu dostawczego",
            "834201 — Operator wózka widłowego",
            "941201 — Pomoc kuchenna",
            "512001 — Kucharz",
            "751201 — Cukiernik",
            "753303 — Krawiec",
            "721301 — Blacharz",
            "721205 — Spawacz",
            "711101 — Murarz",
            "711301 — Cieśla",
            "711401 — Zbrojarz",
            "712601 — Hydraulik",
            "741101 — Elektryk",
            "816003 — Operator maszyn produkcyjnych",
            "911207 — Sprzątaczka",
            "912101 — Praczka",
            "524902 — Sprzedawca",
            "331301 — Księgowy",
            "242217 — Specjalista ds. kadr",
            "143990 — Kierownik (gdzie indziej niesklasyfikowany)",
            "522301 — Kasjer",
            "432103 — Magazynier",
            "962902 — Pracownik ochrony",
        ]
        kod_zawodu_select = st.selectbox("Kod zawodu", KODY_ZAWODOW, help="Wybierz z listy lub wpisz poniżej własny")
        kod_zawodu_custom = st.text_input("Lub wpisz własny kod (6 cyfr)", max_chars=6, placeholder="np. 932101")
        kod_zawodu_input = kod_zawodu_custom.strip() if kod_zawodu_custom.strip() else kod_zawodu_select.split(" — ")[0].strip()
        kod_tytulu_val = kod_tytulu_input.strip().zfill(4)

        if typ_operacji == "ZUA":
            st.caption("Ubezpieczenia:")
            ubezp_em = st.checkbox("Emerytalne", value=True)
            ubezp_rent = st.checkbox("Rentowe", value=True)
            ubezp_chor = st.checkbox("Chorobowe", value=True)
            ubezp_wyp = st.checkbox("Wypadkowe", value=True)
            ubezp_zdr = st.checkbox("Zdrowotne", value=True)
        else:
            ubezp_em = ubezp_rent = ubezp_chor = ubezp_wyp = False
            ubezp_zdr = True

    elif typ_operacji == "ZWUA":
        data_wyrejestrowania = st.date_input("Data wyrejestrowania", value=date.today())
        kod_tytulu_input = st.text_input("Kod tytułu ubezpieczenia (4 cyfry)", value="0110", max_chars=4, help="Wpisz 4-cyfrowy kod, np. 0110, 0411, 2250")
        with st.expander("📖 Podpowiedź — popularne kody"):
            for grupa, opis in KODY_TYTULU_REF.items():
                st.caption(f"**{grupa}**: {opis}")
        kod_tytulu_val = kod_tytulu_input.strip().zfill(4)
        kod_wyrejestrowania = st.selectbox("Kod przyczyny wyrejestrowania", [
            "100 — Ustanie tytułu do ubezpieczeń",
            "500 — Zgon ubezpieczonego",
            "600 — Inna przyczyna",
        ])
        kod_wyr_val = kod_wyrejestrowania.split(" — ")[0]

    elif typ_operacji == "RCA":
        kod_tytulu_input = st.text_input("Kod tytułu ubezpieczenia (4 cyfry)", value="0110", max_chars=4, help="Wpisz 4-cyfrowy kod, np. 0110, 0411, 2250")
        with st.expander("📖 Podpowiedź — popularne kody"):
            for grupa, opis in KODY_TYTULU_REF.items():
                st.caption(f"**{grupa}**: {opis}")
        kod_tytulu_val = kod_tytulu_input.strip().zfill(4)

    elif typ_operacji == "ZIUA":
        st.info("ZIUA: system automatycznie wykryje pracowników ze zmianą paszport → PESEL")

    st.divider()

    # --- DANE PŁATNIKA ---
    st.header("🏢 Dane płatnika")
    nip = st.text_input("NIP", placeholder="", help="Wpisz NIP i kliknij Pobierz dane")
    
    if st.button("🔍 Pobierz dane z GUS/KRS", use_container_width=True, disabled=not nip.strip()):
        with st.spinner("Szukam..."):
            wynik = pobierz_dane_z_nip(nip)
            if wynik:
                st.session_state["auto_nazwa"] = wynik["nazwa"]
                st.session_state["auto_regon"] = wynik["regon"]
                st.session_state["auto_adres"] = wynik.get("adres", "")
                st.session_state["auto_zrodlo"] = wynik["zrodlo"]
                st.success(f"✅ Znaleziono: {wynik['nazwa']}")
                st.caption(f"Źródło: {wynik['zrodlo']}")
                st.rerun()
            else:
                st.error("Nie znaleziono firmy o podanym NIP")
    
    regon = st.text_input("REGON", value=st.session_state.get("auto_regon", ""), placeholder="")
    nazwa = st.text_input("Nazwa firmy (pełna)", value=st.session_state.get("auto_nazwa", ""), placeholder="")
    nazwa_skr = st.text_input("Nazwa skrócona (max 31 znaków)", placeholder="", max_chars=31, help="Wymagana dla spółek — max 31 znaków w KEDU")
    if nazwa and not nazwa_skr and len(nazwa) > 31:
        st.warning(f"⚠️ Nazwa ma {len(nazwa)} znaków — wpisz skróconą (max 31)")
    pesel_pl = st.text_input("PESEL właściciela", placeholder="")
    nazwisko_pl = st.text_input("Nazwisko właściciela", placeholder="")
    imie_pl = st.text_input("Imię właściciela", placeholder="")
    data_ur_pl = st.text_input("Data urodzenia właśc.", placeholder="")

    if nip and len(nip.strip()) != 10:
        st.warning("NIP powinien mieć 10 cyfr")

    st.divider()

    # --- OKRES ---
    st.header("📅 Okres")
    col1, col2 = st.columns(2)
    rok = col1.number_input("Rok", value=2026, min_value=2020, max_value=2030)
    miesiac = col2.number_input("Miesiąc", value=datetime.now().month, min_value=1, max_value=12)

    st.divider()

    # --- BIURO ---
    st.header("🖨️ Dane biura (na PDF)")
    nazwa_biura = st.text_input("Nazwa biura", value="")
    adres_biura = st.text_input("Adres biura", placeholder="")

    st.divider()
    st.header("🔑 API Anthropic")
    api_key = st.text_input("Klucz API", type="password", help="Do skanowania dokumentów (Claude Vision)")
    st.divider()
    st.header("📊 Google Sheets")
    sheets_id = st.text_input("ID arkusza", value="1gRvYs6CaTpQpz717YI1etNqqJItEaYnBEqmNMPJZBEU", help="ID z URL arkusza Google Sheets")
    sheets_gid = st.number_input("GID arkusza", value=18649467, help="GID z URL (po #gid=)")
    sheets_creds = st.text_area("Klucz JSON Service Account", height=100, help="Wklej cały JSON klucza service account z Google Cloud", placeholder='{"type":"service_account","project_id":"...","private_key":"..."}')




def _ustaw_config():
    config.NIP_PLATNIKA = nip.strip()
    config.REGON_PLATNIKA = regon.strip()
    config.NAZWA_PLATNIKA = nazwa.strip()
    config.NAZWA_SKROCONA = nazwa_skr.strip()
    config.PESEL_PLATNIKA = pesel_pl.strip()
    config.NAZWISKO_PLATNIKA = nazwisko_pl.strip()
    config.IMIE_PLATNIKA = imie_pl.strip()
    config.DATA_UR_PLATNIKA = data_ur_pl.strip()
    config.OKRES_ROK = int(rok)
    config.OKRES_MIESIAC = int(miesiac)
    config.NAZWA_BIURA = nazwa_biura
    config.ADRES_BIURA = adres_biura


# ============================================================
#  UPLOAD + EDYCJA
# ============================================================

# ============================================================
#  SKANER DOKUMENTÓW
# ============================================================
with st.expander("📷 Skaner dokumentów — odczyt z paszportu / powiadomienia PESEL", expanded=False):
    st.caption("Wgraj zdjęcie paszportu lub powiadomienia o nadaniu PESEL — system odczyta dane automatycznie")
    
    col_scan1, col_scan2 = st.columns([2, 1])
    with col_scan1:
        scan_file = st.file_uploader("Wgraj dokument", type=["jpg", "jpeg", "png", "pdf", "webp"], key="scanner")
    with col_scan2:
        scan_typ = st.radio("Typ dokumentu", ["Auto", "Paszport", "PESEL"], horizontal=True)
    
    if scan_file and st.button("🔍 Skanuj dokument", type="primary"):
        if not api_key:
            st.error("Wpisz klucz API Anthropic w panelu po lewej (sekcja 🔑)")
        else:
            with st.spinner("Analizuję dokument..."):
                try:
                    typ_map = {"Auto": "auto", "Paszport": "paszport", "PESEL": "pesel"}
                    wynik = skanuj_dokument(
                        image_bytes=scan_file.read(),
                        typ=typ_map[scan_typ],
                        api_key=api_key,
                    )
                    
                    if wynik.imie or wynik.nazwisko:
                        st.success(f"✅ Odczytano: **{wynik.imie} {wynik.nazwisko}** ({wynik.typ_dokumentu}, pewność: {wynik.pewnosc})")
                        
                        col_r1, col_r2 = st.columns(2)
                        with col_r1:
                            st.text_input("Imię", value=wynik.imie, key="scan_imie", disabled=True)
                            st.text_input("Nazwisko", value=wynik.nazwisko, key="scan_nazwisko", disabled=True)
                            st.text_input("Drugie imię", value=wynik.drugie_imie, key="scan_drugie", disabled=True)
                            st.text_input("PESEL", value=wynik.pesel, key="scan_pesel", disabled=True)
                        with col_r2:
                            st.text_input("Nr dokumentu", value=wynik.nr_dokumentu, key="scan_nr", disabled=True)
                            st.text_input("Obywatelstwo", value=wynik.obywatelstwo, key="scan_obyw", disabled=True)
                            st.text_input("Data urodzenia", value=wynik.data_urodzenia, key="scan_data", disabled=True)
                            st.text_input("Płeć", value=wynik.plec, key="scan_plec", disabled=True)
                        
                        if wynik.uwagi:
                            st.info(f"📝 Uwagi: {wynik.uwagi}")
                        if wynik.mrz:
                            with st.expander("MRZ"):
                                st.code(wynik.mrz)
                        
                        # Kopiowanie do schowka
                        dane_csv = f"{wynik.nazwisko}\t{wynik.imie}\t{wynik.drugie_imie}\t{wynik.pesel}\t{wynik.nr_dokumentu}\t{wynik.obywatelstwo}\t{wynik.data_urodzenia}\t{wynik.plec}"
                        st.code(dane_csv, language=None)
                        st.caption("↑ Skopiuj i wklej do Excela (Tab-separated)")
                    else:
                        st.error("Nie udało się odczytać danych z dokumentu")
                        if wynik.uwagi:
                            st.info(wynik.uwagi)
                except Exception as e:
                    st.error(f"Błąd skanowania: {e}")

st.divider()

uploaded = st.file_uploader(
    "📁 Wrzuć plik Excel z danymi pracowników",
    type=["xlsx", "xls", "csv"],
)

if uploaded:
    with tempfile.NamedTemporaryFile(suffix=Path(uploaded.name).suffix, delete=False) as tmp:
        tmp.write(uploaded.read())
        tmp_path = tmp.name

    try:
        pracownicy = wczytaj_excel(tmp_path)
    except Exception as e:
        st.error(f"Błąd wczytywania: {e}")
        pracownicy = []

    if pracownicy:
        # Buduj DataFrame do edycji
        rows = []
        for p in pracownicy:
            rows.append({
                "✓": True,  # checkbox — czy uwzględnić
                "Nazwisko": p.nazwisko,
                "Imię": p.imie,
                "PESEL": p.pesel or "",
                "Nr paszportu": p.nr_paszportu or "",
                "Data urodzenia": str(p.data_urodzenia)[:10] if p.data_urodzenia else "",
                "Obywatelstwo": p.obywatelstwo or "PL",
                "Kod pocztowy": p.kod_pocztowy or "",
                "Miejscowość": p.miejscowosc or "",
                "Ulica": p.ulica or "",
                "Nr domu": p.nr_domu or "",
                "Nr lokalu": p.nr_lokalu or "",
                "Kod tytułu": p.kod_tytulu or "0110",
                "Kod zawodu": p.kod_zawodu or "",
                "NFZ (auto)": p.kod_nfz or kod_pocztowy_do_nfz(p.kod_pocztowy) or "",
                "Data zgłoszenia": str(p.data_zgłoszenia)[:10] if p.data_zgłoszenia else "",
            })

        df = pd.DataFrame(rows)

        # === STATYSTYKI ===
        st.markdown(f"### 📊 Wczytano {len(df)} pracowników — operacja: **{operacja}**")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Łącznie", len(df))
        col2.metric("Z PESEL", sum(1 for _, r in df.iterrows() if r["PESEL"]))
        col3.metric("Z paszportem", sum(1 for _, r in df.iterrows() if r["Nr paszportu"]))
        obcokr = sum(1 for _, r in df.iterrows() if r["Obywatelstwo"] not in ("PL", "POLSKIE", ""))
        col4.metric("Obcokrajowcy", obcokr)

        # === EDYTOWALNA TABELA ===
        st.markdown("### ✏️ Edycja danych")
        st.caption("Odznacz checkbox żeby pominąć pracownika. Edytuj dowolne pole przed generowaniem.")

        # Nadpisz daty i kody z sidebara (jeśli ustawione)
        if typ_operacji in ("ZUA", "ZZA"):
            df["Kod tytułu"] = kod_tytulu_val
            df["Kod zawodu"] = kod_zawodu_input.strip()
            df["Data zgłoszenia"] = str(data_zgloszenia)
        elif typ_operacji == "ZWUA":
            df["Kod tytułu"] = kod_tytulu_val
            df["Kod zawodu"] = kod_zawodu_input.strip() if "kod_zawodu_input" in dir() else ""
            if "Data wyrejestrowania" not in df.columns:
                df["Data wyrejestrowania"] = str(data_wyrejestrowania)

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",  # pozwala dodawać/usuwać wiersze
            column_config={
                "✓": st.column_config.CheckboxColumn("✓", default=True, width="small"),
                "PESEL": st.column_config.TextColumn("PESEL", width="medium"),
                "Nr paszportu": st.column_config.TextColumn("Nr paszportu", width="medium"),
                "Kod tytułu": st.column_config.TextColumn("Kod tytułu", width="small"),
                "NFZ (auto)": st.column_config.SelectboxColumn("NFZ",
                    options=[f"{i:02d}" for i in range(1, 17)], width="small"),
                "Obywatelstwo": st.column_config.TextColumn("Obywatelstwo", width="small"),
            },
            hide_index=True,
        )

        # Filtruj zaznaczonych
        selected = edited_df[edited_df["✓"] == True] if "✓" in edited_df.columns else edited_df
        st.info(f"Zaznaczonych do przetworzenia: **{len(selected)}** z {len(edited_df)}")

        # === KONWERSJA Z POWROTEM NA PRACOWNIKÓW ===
        def df_do_pracownikow(df_sel):
            from excel_reader import Pracownik, waliduj_pracownika
            result = []
            for _, r in df_sel.iterrows():
                p = Pracownik(
                    pesel=str(r.get("PESEL", "")).strip(),
                    nazwisko=str(r.get("Nazwisko", "")).strip(),
                    imie=str(r.get("Imię", "")).strip(),
                    data_urodzenia=str(r.get("Data urodzenia", "")).strip(),
                    kod_pocztowy=str(r.get("Kod pocztowy", "")).strip(),
                    miejscowosc=str(r.get("Miejscowość", "")).strip(),
                    ulica=str(r.get("Ulica", "")).strip(),
                    nr_domu=str(r.get("Nr domu", "")).strip(),
                    nr_lokalu=str(r.get("Nr lokalu", "")).strip(),
                    kod_tytulu=str(r.get("Kod tytułu", "0110")).strip(),
                    kod_zawodu=str(r.get("Kod zawodu", "")).strip(),
                    kod_nfz=str(r.get("NFZ (auto)", "")).strip(),
                    nr_paszportu=str(r.get("Nr paszportu", "")).strip(),
                    obywatelstwo=str(r.get("Obywatelstwo", "PL")).strip(),
                    typ_deklaracji=typ_operacji,
                    data_zgłoszenia=str(r.get("Data zgłoszenia", "")).strip(),
                )
                # Auto-NFZ
                if not p.kod_nfz and p.kod_pocztowy:
                    p.kod_nfz = kod_pocztowy_do_nfz(p.kod_pocztowy)
                # Obcokrajowiec
                if p.obywatelstwo and p.obywatelstwo.upper() not in ("PL", "POLSKIE", "POLAND", ""):
                    p.czy_obcokrajowiec = True
                if p.nr_paszportu:
                    p.typ_identyfikatora = "2"
                    p.czy_obcokrajowiec = True
                # Ubezpieczenia (z sidebara)
                if typ_operacji == "ZUA":
                    p.ubezp_emerytalne = ubezp_em
                    p.ubezp_rentowe = ubezp_rent
                    p.ubezp_chorobowe = ubezp_chor
                    p.ubezp_wypadkowe = ubezp_wyp
                    p.ubezp_zdrowotne = ubezp_zdr
                elif typ_operacji == "ZZA":
                    p.ubezp_emerytalne = False
                    p.ubezp_rentowe = False
                    p.ubezp_chorobowe = False
                    p.ubezp_wypadkowe = False
                    p.ubezp_zdrowotne = True

                p.bledy = waliduj_pracownika(p)
                result.append(p)
            return result

        # === GENEROWANIE ===
        st.divider()
        st.markdown("### 🚀 Generuj")

        if not nip.strip():
            st.warning("⚠️ Wpisz **NIP płatnika** w panelu po lewej stronie!")
        else:
            col_kedu, col_pdf = st.columns(2)

            with col_kedu:
                st.subheader("📄 KEDU XML")
                if st.button("🔨 Generuj KEDU", type="primary", use_container_width=True):
                    _ustaw_config()
                    prac = df_do_pracownikow(selected)
                    poprawni = [p for p in prac if not p.bledy]
                    bledni = [p for p in prac if p.bledy]

                    if bledni:
                        with st.expander(f"⚠️ {len(bledni)} rekordów z błędami (pominięte)", expanded=True):
                            for p in bledni:
                                st.error(f"**{p.nazwisko} {p.imie}**: {', '.join(p.bledy)}")

                    if poprawni:
                        with tempfile.TemporaryDirectory() as tmpdir:
                            try:
                                wyniki = generuj_wszystkie_kedu(poprawni, tmpdir)
                                if wyniki:
                                    if len(wyniki) == 1:
                                        sciezka = list(wyniki.values())[0]
                                        st.session_state.kedu_bytes = sciezka.read_bytes()
                                        st.session_state.kedu_filename = sciezka.name
                                    else:
                                        buf = io.BytesIO()
                                        with zipfile.ZipFile(buf, 'w') as zf:
                                            for typ, sciezka in wyniki.items():
                                                zf.write(sciezka, sciezka.name)
                                        st.session_state.kedu_bytes = buf.getvalue()
                                        st.session_state.kedu_filename = f"KEDU_{rok}_{miesiac:02d}.zip"
                                    for typ, sciezka in wyniki.items():
                                        st.success(f"✅ {typ}: {sciezka.name} ({len(poprawni)} pracowników)")
                                else:
                                    st.error("Nie wygenerowano plików — sprawdź typ operacji")
                            except ValueError as e:
                                st.error(str(e))
                            except Exception as e:
                                st.error(f"Błąd: {e}")
                    else:
                        st.error("Brak poprawnych rekordów")

                if st.session_state.kedu_bytes:
                    st.download_button(
                        "⬇️ Pobierz KEDU",
                        data=st.session_state.kedu_bytes,
                        file_name=st.session_state.kedu_filename,
                        mime="application/xml" if st.session_state.kedu_filename.endswith(".xml") else "application/zip",
                        use_container_width=True,
                    )

            with col_pdf:
                st.subheader("📋 Potwierdzenia PDF")
                if st.button("🔨 Generuj PDF-y", type="primary", use_container_width=True):
                    _ustaw_config()
                    prac = df_do_pracownikow(selected)
                    poprawni = [p for p in prac if not p.bledy]

                    if poprawni:
                        with tempfile.TemporaryDirectory() as tmpdir:
                            try:
                                prefix = f"ZUS-{rok}-{miesiac:02d}-"
                                sciezki = generuj_potwierdzenia(poprawni, tmpdir, nr_potwierdzenia_prefix=prefix)
                                if sciezki:
                                    buf = io.BytesIO()
                                    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                                        for s in sciezki:
                                            zf.write(s, s.name)
                                    st.session_state.pdf_zip_bytes = buf.getvalue()
                                    st.success(f"✅ {len(sciezki)} potwierdzeń PDF")
                            except Exception as e:
                                st.error(f"Błąd: {e}")
                    else:
                        st.error("Brak poprawnych rekordów")

                if st.session_state.pdf_zip_bytes:
                    st.download_button(
                        "⬇️ Pobierz PDF-y (ZIP)",
                        data=st.session_state.pdf_zip_bytes,
                        file_name=f"potwierdzenia_{rok}_{miesiac:02d}.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )

            # Podgląd KEDU
            if st.session_state.kedu_bytes and st.session_state.kedu_filename.endswith(".xml"):
                with st.expander("👁 Podgląd KEDU XML"):
                    st.code(st.session_state.kedu_bytes.decode("utf-8")[:5000], language="xml")

# === FOOTER ===
st.divider()
st.caption("Płatnik Auto v3.0 — ABACUS CENTRUM KSIĘGOWE")
