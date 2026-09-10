#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit — scraper cennika triverna.pl (port interfejsu z triverna_gui.py).

Rdzen logiki (triverna_scraper.py, triverna_xlsx.py, mandala_fetcher.py) jest
w pelni wspoldzielony z wersja desktopowa (.exe) - ten plik to WYLACZNIE
warstwa interfejsu.

Auto-pobieranie z Mandali dziala TYLKO gdy aplikacja ma dostep do
wewnetrznej sieci firmowej i lokalnie zainstalowanego Chrome - w innych
srodowiskach (np. Streamlit Community Cloud, uzywane na czas testow) ta
opcja jest automatycznie ukrywana, a uzytkownik moze zamiast tego wgrac
plik z obecnym cennikiem recznie. Po przeniesieniu aplikacji na docelowy
serwer (ten sam, na ktorym dziala Mandala) auto-pobieranie zacznie
dzialac samo, bez zadnych zmian w kodzie.

UWAGA (repo publiczne): adres Mandali NIE jest zaszyty w kodzie - jest
odczytywany z st.secrets["mandala_url"] (lokalnie: .streamlit/secrets.toml,
na Streamlit Cloud: panel ustawien aplikacji -> Secrets). Bez skonfigurowanego
sekretu auto-pobieranie jest po prostu niedostepne (dokladnie jak wtedy, gdy
srodowisko nie ma dostepu do sieci wewnetrznej) - nigdy nie wywala aplikacji.
"""

import os
import tempfile
from datetime import date, timedelta

import streamlit as st

import triverna_scraper as core
import triverna_xlsx as xl

try:
    import mandala_fetcher as mf
    _MANDALA_IMPORT_OK = True
except Exception:
    mf = None
    _MANDALA_IMPORT_OK = False


st.set_page_config(page_title="Triverna Scraper", page_icon="👑", layout="centered")


def _get_mandala_url() -> str | None:
    """Adres Mandali NIGDY nie jest zaszyty w kodzie (repo jest publiczne) -
    tylko st.secrets albo zmienna srodowiskowa, w tej kolejnosci. Brak
    obu = auto-pobieranie po prostu niedostepne (nie blad)."""
    try:
        val = st.secrets.get("mandala_url")
        if val:
            return val
    except Exception:
        pass
    return os.environ.get("MANDALA_URL")


@st.cache_data(ttl=30, show_spinner=False)
def _mandala_reachable(mandala_url: str | None, timeout: float = 2.0) -> bool:
    """Szybki, tani test - czy w tym srodowisku widac wewnetrzny serwer
    Mandali w ogole (bez odpalania Playwrighta/Chrome). Wynik cache'owany
    na 30s, zeby nie sprawdzac tego przy kazdym przerysowaniu strony."""
    if not _MANDALA_IMPORT_OK or not mandala_url:
        return False
    try:
        import requests
        requests.head(mandala_url, timeout=timeout)
        return True
    except Exception:
        return False


def _tmp_dir() -> str:
    d = os.path.join(tempfile.gettempdir(), "triverna_streamlit")
    os.makedirs(d, exist_ok=True)
    return d


st.title("👑 Triverna — cennik konkurencji")
st.caption("Wklej link do oferty, wybierz zakres dat i wygeneruj cennik w formacie Mandala.")

mandala_url = _get_mandala_url()
mandala_ok = _mandala_reachable(mandala_url)


# UWAGA: celowo BEZ st.form() - formularze w Streamlicie nie przerysowuja
# sie (nie pokazuja/chowaja warunkowych widgetow, np. uploadera pliku)
# dopoki caly formularz nie zostanie wyslany, wiec radio->pokaz-uploader
# nie dzialaloby "na zywo" w jego srodku. Zwykle, "plaskie" widgety
# przerysowuja strone przy kazdej zmianie, co daje reaktywny interfejs.
url = st.text_input(
    "Link do oferty na triverna.pl",
    placeholder="https://triverna.pl/hotel/nazwa-hotelu",
)

col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input(
        "Od (data przyjazdu)", value=date.today() + timedelta(days=1), format="DD/MM/YYYY"
    )
with col2:
    end_date = st.date_input(
        "Do (data przyjazdu)", value=date.today() + timedelta(days=180), format="DD/MM/YYYY"
    )

with st.expander("Obłożenie (domyślnie 2 dorosłych)"):
    c1, c2, c3, c4 = st.columns(4)
    adults = c1.number_input("Dorośli", min_value=1, max_value=8, value=2)
    children = c2.number_input("Dzieci", min_value=0, max_value=6, value=0)
    babies = c3.number_input("Niemowlęta", min_value=0, max_value=4, value=0)
    rooms_booked = c4.number_input("Liczba pokoi", min_value=1, max_value=5, value=1)

with st.expander("Szybkość scrapowania (zaawansowane)"):
    max_workers = st.slider(
        "Równoległe wątki pobierania cen", min_value=1, max_value=8, value=1,
        help=(
            "1 = dotychczasowe, sekwencyjne działanie (bezpieczny wybór). Wyższe wartości "
            "przyspieszają scrapowanie, ale zwiększają liczbę równoległych połączeń do "
            "triverna.pl — podnoś stopniowo i sprawdzaj na małym zakresie dat, zanim użyjesz "
            "większej wartości na produkcji."
        ),
    )

report_choice = st.radio(
    "Co wygenerować?",
    ["Cennik Mandala", "Cennik Mandala + raport dostępności (Ava)"],
    horizontal=True,
)

st.divider()
st.subheader("Obecny cennik (opcjonalnie)")

base_options = ["Nie scalaj — świeży plik od zera"]
if mandala_ok:
    base_options.append("Auto-pobierz z Mandali")
base_options.append("Wgraj plik ręcznie")
if not mandala_ok:
    st.caption(
        "🔌 Auto-pobieranie z Mandali niedostępne w tym środowisku (brak dostępu do "
        "sieci wewnętrznej firmy) — wgraj plik ręcznie albo pomiń scalanie. "
        "Po uruchomieniu na docelowym serwerze ta opcja pojawi się automatycznie."
    )

base_mode = st.radio("Źródło obecnych cen", base_options)

uploaded_file = None
if base_mode == "Wgraj plik ręcznie":
    uploaded_file = st.file_uploader("Plik z obecnym cennikiem (.xlsx, format Mandala)", type=["xlsx"])

best_price_only = False
diff_report = False
if base_mode != "Nie scalaj — świeży plik od zera":
    best_price_only = st.checkbox(
        "Najlepsza cena (nadpisuj tylko gdy konkurencja jest tańsza od obecnej)"
    )
    if best_price_only:
        diff_report = st.checkbox("Dodatkowy raport różnic (co się zmieniło)")

show_log = st.checkbox("Pokaż szczegóły techniczne", value=False)

submitted = st.button("🚀 Uruchom scrapowanie", use_container_width=True, type="primary")


if submitted:
    if not url.strip():
        st.error("Podaj link do oferty.")
        st.stop()
    if end_date <= start_date:
        st.error("Data końcowa musi być późniejsza niż początkowa.")
        st.stop()
    if base_mode == "Wgraj plik ręcznie" and uploaded_file is None:
        st.error("Wybrałeś ręczny upload, ale nie wgrałeś pliku.")
        st.stop()

    log_lines = []

    def log(msg):
        log_lines.append(str(msg))

    progress_bar = st.progress(0.0, text="Przygotowuję...")

    def progress_cb(idx, total, day_str):
        pct = min(idx / total, 1.0) if total else 0.0
        progress_bar.progress(pct, text=f"Sprawdzam {day_str}... ({idx}/{total})")

    tmp_dir = _tmp_dir()
    base_file_path = None

    try:
        if base_mode == "Wgraj plik ręcznie":
            base_file_path = os.path.join(tmp_dir, "uploaded_" + uploaded_file.name)
            with open(base_file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

        with st.spinner("Scrapuję ceny z Triverny... (może potrwać kilka minut)"):
            rows, hotel_name, calendar_days = core.scrape(
                url=url.strip(),
                start_date=start_date,
                end_date=end_date,
                adults=int(adults),
                children=int(children),
                babies=int(babies),
                rooms_booked=int(rooms_booked),
                fixed_nights=None,
                delay=0.15,
                log=log,
                progress_callback=progress_cb,
                max_workers=int(max_workers),
            )

        progress_bar.progress(1.0, text="Scrapowanie zakończone.")

        if not rows:
            st.warning(
                "Brak jakichkolwiek dostępnych dat przyjazdu w wybranym zakresie — "
                "sprawdź inny zakres dat lub link do oferty."
            )
            st.stop()

        if base_mode == "Auto-pobierz z Mandali":
            with st.spinner("Pobieram obecny cennik z Mandali..."):
                mandala_result = mf.fetch_current_prices(hotel_name, tmp_dir, mandala_url=mandala_url, log=log)
                base_file_path = mandala_result.file_path
                if mandala_result.warning:
                    st.warning(mandala_result.warning)

        run_date = date.today()
        outputs = []

        if base_file_path:
            out_name = core.build_output_filename(hotel_name, run_date, suffix="_mandala", ext="xlsx")
            out_path = os.path.join(tmp_dir, out_name)
            merge_result = xl.write_mandala_xlsx_merge(
                rows, calendar_days, base_file_path, out_path,
                best_price_only=best_price_only, log=log,
            )
            outputs.append(("Cennik Mandala (scalony)", out_path))

            matched, total = merge_result["matched_rooms"], merge_result["total_rooms"]
            if total and matched < total:
                st.warning(
                    f"Dopasowano {matched}/{total} typów pokoi do przesłanego pliku — "
                    f"pozostałe pokoje: {', '.join(merge_result['unmatched_names'])} "
                    "nie zostały zaktualizowane (sprawdź nazewnictwo ręcznie)."
                )
            elif total:
                st.success(f"Dopasowano wszystkie {total} typy pokoi do przesłanego pliku.")

            if best_price_only:
                changes = merge_result.get("changes") or []
                if diff_report:
                    diff_name = core.build_output_filename(hotel_name, run_date, suffix="_zmiany", ext="xlsx")
                    diff_path = os.path.join(tmp_dir, diff_name)
                    xl.write_price_diff_report(changes, hotel_name, diff_path, log=log)
                    outputs.append(("Raport zmian cen", diff_path))
                if not changes:
                    st.info("Ceny 1:1 — konkurencja nie była tańsza w żadnym miejscu, nic nie zmieniono.")
                else:
                    st.success(f"Znaleziono {len(changes)} cen tańszych u konkurencji — zaktualizowano.")
        else:
            out_name = core.build_output_filename(hotel_name, run_date, suffix="_mandala", ext="xlsx")
            out_path = os.path.join(tmp_dir, out_name)
            xl.write_mandala_xlsx(rows, calendar_days, out_path, log=log)
            outputs.append(("Cennik Mandala", out_path))

        if report_choice.endswith("(Ava)"):
            ava_name = core.build_output_filename(hotel_name, run_date, suffix="_ava", ext="xlsx")
            ava_path = os.path.join(tmp_dir, ava_name)
            xl.write_ava_xlsx(rows, ava_path, log=log)
            outputs.append(("Raport dostępności (Ava)", ava_path))

        st.subheader("📥 Gotowe pliki")
        for label, path in outputs:
            with open(path, "rb") as f:
                st.download_button(
                    f"Pobierz: {label}",
                    data=f.read(),
                    file_name=os.path.basename(path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

    except core.ScrapeCancelled:
        st.info("Przerwano.")
    except Exception as exc:
        st.error(f"Błąd: {exc}")
        log(f"[WYJĄTEK] {type(exc).__name__}: {exc}")

    if show_log:
        with st.expander("Szczegóły techniczne", expanded=True):
            st.code("\n".join(log_lines) or "(brak logów)", language=None)
