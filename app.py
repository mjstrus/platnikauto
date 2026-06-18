"""
Płatnik Auto — Streamlit App
Wrzuć Excel + wpisz dane płatnika → KEDU XML + potwierdzenia PDF

Uruchomienie:
  pip install streamlit openpyxl pandas fpdf2
  streamlit run app.py
"""
import streamlit as st
import pandas as pd
import tempfile
import zipfile
import io
import sys
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))

import config
from excel_reader import wczytaj_excel, raport_walidacji, kod_pocztowy_do_nfz
from kedu_generator import generuj_wszystkie_kedu
from pdf_confirmations import generuj_potwierdzenia

# === PAGE CONFIG ===
st.set_page_config(
    page_title="Płatnik Auto",
    page_icon="📋",
    layout="wide",
)

# === STYLE ===
st.markdown("""
<style>
    .block-container { padding-top: 2rem; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { padding: 8px 20px; }
    div[data-testid="stMetric"] { background: #0f172a; padding: 12px 16px; border-radius: 8px; border-left: 3px solid #3b82f6; }
</style>
""", unsafe_allow_html=True)

# === HEADER ===
st.markdown("# 📋 Płatnik Auto")
st.markdown("*Wrzuć Excel → wpisz dane płatnika → pobierz KEDU + potwierdzenia PDF*")
st.divider()

# === SESSION STATE ===
if "pracownicy" not in st.session_state:
    st.session_state.pracownicy = []
if "kedu_bytes" not in st.session_state:
    st.session_state.kedu_bytes = None
if "pdf_zip_bytes" not in st.session_state:
    st.session_state.pdf_zip_bytes = None
if "kedu_filename" not in st.session_state:
    st.session_state.kedu_filename = ""
if "processed" not in st.session_state:
    st.session_state.processed = False

# === SIDEBAR — DANE PŁATNIKA ===
with st.sidebar:
    st.header("⚙️ Dane płatnika")
    st.caption("Uzupełnij dla firmy, dla której generujesz deklaracje")

    nip = st.text_input("NIP", placeholder="7161212863", help="10 cyfr, bez kresek")
    regon = st.text_input("REGON", placeholder="380505464")
    nazwa = st.text_input("Nazwa firmy", placeholder='"U FRYZJERA"')
    pesel_pl = st.text_input("PESEL właściciela", placeholder="71111104448", help="Dla osób fizycznych")
    nazwisko_pl = st.text_input("Nazwisko właściciela", placeholder="MATYJASZEK")
    imie_pl = st.text_input("Imię właściciela", placeholder="AGNIESZKA")
    data_ur_pl = st.text_input("Data urodzenia właśc.", placeholder="1971-11-11")

    st.divider()
    st.subheader("📅 Okres")
    col1, col2 = st.columns(2)
    rok = col1.number_input("Rok", value=2026, min_value=2020, max_value=2030)
    miesiac = col2.number_input("Miesiąc", value=datetime.now().month, min_value=1, max_value=12)

    st.divider()
    st.subheader("🏢 Dane biura (na PDF-ach)")
    nazwa_biura = st.text_input("Nazwa biura", value="ABACUS CENTRUM KSIĘGOWE")
    adres_biura = st.text_input("Adres biura", placeholder="ul. Przykładowa 1, 24-100 Puławy")

    # Walidacja NIP
    if nip and len(nip.strip()) != 10:
        st.warning("NIP powinien mieć 10 cyfr")


def _ustaw_config():
    """Ustawia config z wartości wpisanych w sidebar"""
    config.NIP_PLATNIKA = nip.strip()
    config.REGON_PLATNIKA = regon.strip()
    config.NAZWA_PLATNIKA = nazwa.strip()
    config.PESEL_PLATNIKA = pesel_pl.strip()
    config.NAZWISKO_PLATNIKA = nazwisko_pl.strip()
    config.IMIE_PLATNIKA = imie_pl.strip()
    config.DATA_UR_PLATNIKA = data_ur_pl.strip()
    config.OKRES_ROK = int(rok)
    config.OKRES_MIESIAC = int(miesiac)
    config.NAZWA_BIURA = nazwa_biura
    config.ADRES_BIURA = adres_biura


# === MAIN — UPLOAD ===
uploaded = st.file_uploader(
    "📁 Wrzuć plik Excel z danymi pracowników",
    type=["xlsx", "xls", "csv"],
    help="Format biurowy (Imię i nazwisko, Typ dokumentu, ...) lub szablon rozszerzony"
)

if uploaded:
    # Zapisz do tempa i wczytaj
    with tempfile.NamedTemporaryFile(suffix=Path(uploaded.name).suffix, delete=False) as tmp:
        tmp.write(uploaded.read())
        tmp_path = tmp.name

    try:
        pracownicy = wczytaj_excel(tmp_path)
        st.session_state.pracownicy = pracownicy
        st.session_state.processed = False
        st.session_state.kedu_bytes = None
        st.session_state.pdf_zip_bytes = None
    except Exception as e:
        st.error(f"Błąd wczytywania: {e}")
        pracownicy = []

    if pracownicy:
        poprawni = [p for p in pracownicy if not p.bledy]
        bledni = [p for p in pracownicy if p.bledy]

        # === STATYSTYKI ===
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Łącznie", len(pracownicy))
        col2.metric("Poprawne", len(poprawni))
        col3.metric("Błędy", len(bledni))
        rca_count = sum(1 for p in poprawni if p.typ_deklaracji == "RCA")
        zua_count = sum(1 for p in poprawni if p.typ_deklaracji == "ZUA")
        col4.metric("RCA", rca_count)
        col5.metric("ZUA", zua_count)

        # === TABKI ===
        tab_dane, tab_bledy, tab_generuj = st.tabs(["📊 Dane pracowników", "⚠️ Błędy", "🚀 Generuj"])

        with tab_dane:
            # Tabela danych
            df_data = []
            for p in pracownicy:
                status = "✅" if not p.bledy else "❌"
                df_data.append({
                    "Status": status,
                    "Nazwisko": p.nazwisko,
                    "Imię": p.imie,
                    "PESEL": p.pesel or "—",
                    "Paszport": p.nr_paszportu or "—",
                    "Typ": p.typ_deklaracji,
                    "Kod tytułu": p.kod_tytulu,
                    "NFZ": p.kod_nfz or "auto",
                    "Obywatelstwo": p.obywatelstwo or "PL",
                    "Data zgł.": str(p.data_zgłoszenia)[:10] if p.data_zgłoszenia else "—",
                    "Adres": f"{p.kod_pocztowy} {p.miejscowosc}" if p.miejscowosc else "—",
                })
            st.dataframe(pd.DataFrame(df_data), use_container_width=True, hide_index=True)

        with tab_bledy:
            if bledni:
                for p in bledni:
                    st.error(f"**{p.nazwisko} {p.imie}** ({p.pesel or p.nr_paszportu}): {', '.join(p.bledy)}")
            else:
                st.success("Brak błędów — wszystkie rekordy poprawne!")

        with tab_generuj:
            if not nip.strip():
                st.warning("⚠️ Wpisz NIP płatnika w panelu po lewej stronie!")
            elif not poprawni:
                st.warning("Brak poprawnych rekordów do wygenerowania")
            else:
                st.info(f"Gotowe do wygenerowania: **{len(poprawni)} pracowników** dla **{nazwa or nip}** za okres **{rok}/{miesiac:02d}**")

                col_kedu, col_pdf = st.columns(2)

                with col_kedu:
                    st.subheader("📄 KEDU XML")
                    if st.button("🔨 Generuj KEDU", type="primary", use_container_width=True):
                        _ustaw_config()
                        with tempfile.TemporaryDirectory() as tmpdir:
                            try:
                                wyniki = generuj_wszystkie_kedu(poprawni, tmpdir)
                                if wyniki:
                                    # Zbierz wszystkie pliki do jednego ZIP jeśli więcej niż 1
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
                                        st.success(f"✅ {typ}: {sciezka.name}")
                                    st.session_state.processed = True
                                else:
                                    st.error("Nie wygenerowano żadnych plików KEDU")
                            except ValueError as e:
                                st.error(str(e))
                            except Exception as e:
                                st.error(f"Błąd: {e}")

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
                                    st.success(f"✅ Wygenerowano {len(sciezki)} potwierdzeń PDF")
                                else:
                                    st.error("Nie wygenerowano żadnych PDF-ów")
                            except Exception as e:
                                st.error(f"Błąd: {e}")

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
st.caption("Płatnik Auto v2.0 — ABACUS CENTRUM KSIĘGOWE")
