#!/usr/bin/env python3
"""
Automatyczne pobieranie "obecnego cennika" z wewnetrznego systemu Mandala -
Streamlit-owej apki, w ktorej mozna recznie wyszukac hotel i pobrac plik
xlsx z biezacymi cenami z naszego systemu.

W przeciwienstwie do Triverny (czyste zapytania GraphQL/HTTP), Mandala
dziala na sesji Streamlit (WebSocket) - nie da sie tego pobrac prostym
"requests.get()". Zamiast tego uzywamy Playwright do sterowania SYSTEMOWO
zainstalowana przegladarka Chrome (channel="chrome") - NIE pobieramy ani
nie dolaczamy wlasnej przegladarki, wiec wymagany jest zainstalowany
Google Chrome na maszynie, na ktorej dziala program.

Wymaga: `pip install playwright` (NIE trzeba uruchamiac `playwright
install` - channel="chrome" korzysta z systemowej przegladarki, a nie
wbudowanego Chromium pobieranego przez Playwright).

UWAGA: adres Mandali NIE jest tu zaszyty na sztywno (repo jest publiczne) -
podaje go wywolujacy (patrz `mandala_url` w fetch_current_prices) albo
zmienna srodowiskowa MANDALA_URL. W app.py rozwiazywane jest to przez
st.secrets, zeby wartosc nigdy nie trafiala do repozytorium.
"""

import os
import re
import time
from dataclasses import dataclass

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

HOTEL_COMBOBOX_LABEL = "Select hotel(s) for analysis:"
DOWNLOAD_BUTTON_LABEL = "Download Excel"


class MandalaFetchError(Exception):
    """Podnoszone gdy automatyczne pobranie cennika z Mandali sie nie powiedzie."""


@dataclass
class MandalaFetchResult:
    file_path: str
    matched_hotel_label: str
    warning: str | None


def _normalize(name: str) -> str:
    return re.sub(r"\s+", " ", name or "").strip().lower()


def _strip_bracket_id(text: str) -> str:
    """Usuwa koncowy identyfikator rekordu w nawiasie (np. "... [9653]"), zeby
    porownac, czy kilka wpisow w Mandali to ten sam hotel pod ta sama nazwa
    (rozne rekordy), a nie NAPRAWDE rozne hotele."""
    return _normalize(re.sub(r"\s*\[\d+\]\s*$", "", text or ""))


# Ogolne slowa typu pokoju/hotelu, ktore same w sobie nigdy nie sa
# wystarczajaco charakterystyczne, zeby bezpiecznie przeszukiwac po nich
# Mandale (zwracaja zbyt wiele niepowiazanych trafien - sprawdzone na
# realnym przypadku: samo "Hotel" nie zwraca nawet hotelu, ktorego szukamy,
# bo lista wynikow jest ucinana do ~10 pozycji):
_GENERIC_SEARCH_WORDS = {
    "hotel", "resort", "spa", "apartamenty", "apartamenty", "willa",
    "dwor", "palac", "pensjonat", "spa&wellness",
}


def _cascade_search_variants(hotel_name: str) -> list:
    """Zwraca liste zapytan do wyszukania w Mandali, od pelnej (znormalizowanej)
    nazwy hotelu poczynajac, po kolei odcinajac koncowe slowo (np. nazwe
    miasta czy dopisek typu "by Wyndham") - sprawdzone na realnym przypadku:
    Triverna zwraca "Hotel Galaxy Kraków", a w Mandali ten sam hotel figuruje
    jako samo "Hotel Galaxy" (bez miasta). Nie schodzimy do pojedynczego
    ogolnikowego slowa (patrz _GENERIC_SEARCH_WORDS) - zbyt ryzykowne, mogloby
    trafic w zupelnie inny, niepowiazany hotel.
    """
    tokens = _normalize(hotel_name).split()
    variants = []
    for n in range(len(tokens), 0, -1):
        variant_tokens = tokens[:n]
        if n == 1 and variant_tokens[0] in _GENERIC_SEARCH_WORDS:
            continue
        variants.append(" ".join(variant_tokens))
    return variants


def fetch_current_prices(
    hotel_name: str,
    download_dir: str,
    mandala_url: str | None = None,
    log=print,
    headless: bool = True,
    timeout_ms: int = 45000,
) -> MandalaFetchResult:
    """Otwiera Mandale, wyszukuje podanego hotelu po nazwie, generuje raport
    i pobiera plik xlsx z biezacym cennikiem.

    mandala_url: adres Mandali (np. "http://<adres-wewnetrzny>/mandala"). Jesli
    pominiety, brany jest ze zmiennej srodowiskowej MANDALA_URL - nigdy nie
    jest zaszyty na sztywno w kodzie (patrz komentarz na gorze pliku).

    Zwraca MandalaFetchResult (sciezka do pobranego pliku, etykieta
    dopasowanego hotelu w Mandali, oraz ewentualny tekst ostrzezenia
    "UWAGA: ..." ktory Mandala czasem pokazuje dla hoteli-wyjatkow, gdzie
    jej wlasne ceny moga byc bledne).

    Podnosi MandalaFetchError z czytelnym komunikatem, jesli cokolwiek nie
    zadziala (brak adresu, Chrome, sieci wewnetrznej, hotelu, timeout
    generowania).
    """
    mandala_url = mandala_url or os.environ.get("MANDALA_URL")
    if not mandala_url:
        raise MandalaFetchError(
            "Nie skonfigurowano adresu Mandali (brak MANDALA_URL) - w tym srodowisku "
            "automatyczne pobieranie jest niedostepne, wgraj plik recznie."
        )
    os.makedirs(download_dir, exist_ok=True)
    try:
        return _fetch_current_prices_impl(hotel_name, download_dir, mandala_url, log, headless, timeout_ms)
    finally:
        # Proces sterownika Playwrighta (node.exe) fizycznie zyje wewnatrz
        # tymczasowego katalogu, ktory PyInstaller rozpakowuje przy starcie
        # spakowanego .exe (_MEIxxxxxx). Jego zamkniecie (wyzej, przez
        # wyjscie z "with sync_playwright()") jest asynchroniczne na
        # poziomie systemu - proces potrafi zglosic zakonczenie chwile
        # wczesniej, niz Windows faktycznie zwolni uchwyty do jego plikow.
        # Jesli caly program zostanie zamkniety dokladnie w tym waskim
        # oknie (np. zaraz po uzyciu Mandali), bootloader PyInstallera nie
        # zdazy usunac _MEIxxxxxx przy wyjsciu i pokazuje uzytkownikowi
        # okienko "Failed to remove temporary directory" - myslace, nie
        # wplywa na dzialanie programu, ale nie powinno sie pojawiac.
        # Krotkie opoznienie tutaj (raz, po kazdym uzyciu Mandali - nie
        # przy kazdym starcie programu) daje systemowi czas na pelne
        # zwolnienie tych uchwytow, zanim uzytkownik zdazy zamknac program.
        time.sleep(1.5)


def _fetch_current_prices_impl(
    hotel_name: str,
    download_dir: str,
    mandala_url: str,
    log,
    headless: bool,
    timeout_ms: int,
) -> MandalaFetchResult:
    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(channel="chrome", headless=headless)
        except Exception as exc:
            raise MandalaFetchError(
                "Nie udalo sie uruchomic Google Chrome (automatyczne pobieranie z Mandali "
                f"wymaga zainstalowanej przegladarki Chrome na tym komputerze): {exc}"
            ) from exc

        try:
            page = browser.new_page(accept_downloads=True)

            log(f"Mandala: otwieram {mandala_url}...")
            try:
                page.goto(mandala_url, timeout=timeout_ms)
            except PlaywrightTimeoutError as exc:
                raise MandalaFetchError(
                    f"Nie udalo sie otworzyc {mandala_url} - sprawdz czy jestes w sieci "
                    f"wewnetrznej firmy: {exc}"
                ) from exc

            try:
                page.wait_for_selector('[data-testid="stMultiSelect"]', timeout=timeout_ms)
            except PlaywrightTimeoutError as exc:
                raise MandalaFetchError(
                    "Strona Mandali nie wczytala listy hoteli w oczekiwanym czasie"
                ) from exc

            log(f"Mandala: wyszukuje hotel '{hotel_name}'...")
            combo = page.get_by_role("combobox", name=HOTEL_COMBOBOX_LABEL)

            # Probujemy po kolei pelna nazwe, a potem coraz krotsze warianty
            # (odcinajac koncowe slowo - np. miasto). Akceptujemy wynik gdy:
            # - wyszukiwanie zwraca dokladnie JEDNA opcje w ogole (bez wzgledu
            #   na to, czy jej tekst zaczyna sie od zapytania - skoro nie ma
            #   ZADNEGO innego kandydata, nie ma tez czego pomylic; sprawdzone
            #   na realnym przypadku: Magnus Resort figuruje w Mandali jako
            #   "Hotel Magnus Resort", z dodatkowym slowem na poczatku, ktorego
            #   Triverna w ogole nie ma w swojej nazwie - zaden prefiks nigdy
            #   by nie pasowal, mimo ze to jedyny i poprawny wynik), ALBO
            # - kilka opcji zaczyna sie od zapytania, ale wszystkie maja
            #   TAKI SAM tekst poza koncowym identyfikatorem w nawiasie (np.
            #   "Mercure Szczyrk Resort [9653]" i "Mercure Szczyrk Resort
            #   [6222]") - to nie jest niejednoznacznosc miedzy RÓŻNYMI
            #   hotelami, tylko kilka wpisow/rekordow pod ta sama nazwa, wiec
            #   bierzemy pierwszy z nich (tak jak dzialalo to wczesniej).
            # Nigdy nie zgadujemy, gdy zapytanie zwraca kilka NAPRAWDE roznych
            # nazw hoteli - to byloby realne ryzyko cichego pobrania cennika
            # ZUPELNIE INNEGO hotelu.
            chosen = None
            chosen_text = None
            variants = _cascade_search_variants(hotel_name)
            for variant in variants:
                combo.click()
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                combo.type(variant, delay=20)
                page.wait_for_timeout(800)

                options = page.get_by_role("option")
                count = options.count()
                if count == 0:
                    continue

                if count == 1:
                    idx, chosen_text = 0, options.nth(0).inner_text()
                    chosen = options.nth(0)
                    if variant != variants[0]:
                        log(f"  [uwaga] brak dopasowania pelnej nazwy - jedyny wynik po skroconym "
                            f"zapytaniu '{variant}': '{chosen_text}'")
                    break

                all_texts = [options.nth(i).inner_text() for i in range(count)]
                matches = [
                    (i, t) for i, t in enumerate(all_texts) if _normalize(t).startswith(variant)
                ]
                if not matches:
                    continue

                distinct_names = {_strip_bracket_id(t) for _, t in matches}
                if len(distinct_names) == 1:
                    idx, chosen_text = matches[0]
                    chosen = options.nth(idx)
                    if variant != variants[0] or len(matches) > 1:
                        log(f"  [uwaga] '{chosen_text}' - {len(matches)} pasujacych wpisow o tej samej "
                            f"nazwie (rozne rekordy Mandali), biore pierwszy")
                    break
                # Kilka ROZNYCH hoteli pasuje do tego samego zapytania -
                # niejednoznaczne, probujemy krotszy wariant zamiast zgadywac.

            if chosen is None:
                raise MandalaFetchError(
                    f"Mandala nie znalazla JEDNOZNACZNEGO hotelu pasujacego do nazwy '{hotel_name}' "
                    "(sprawdzono pelna nazwe i skrocone warianty) - sprobuj wgrac plik recznie"
                )

            log(f"Mandala: wybieram '{chosen_text}'")
            chosen.click()

            download_btn = page.get_by_role("button", name=DOWNLOAD_BUTTON_LABEL)
            try:
                download_btn.wait_for(state="visible", timeout=timeout_ms)
            except PlaywrightTimeoutError as exc:
                raise MandalaFetchError(
                    f"Mandala nie wygenerowala raportu dla '{chosen_text}' w oczekiwanym czasie"
                ) from exc

            warning = None
            body_text = page.inner_text("body")
            idx = body_text.find("UWAGA")
            if idx != -1:
                warning = body_text[idx : idx + 200].split("\n")[0].strip()
                log(f"  [UWAGA z Mandali] {warning}")

            log("Mandala: pobieram plik Excel...")
            try:
                with page.expect_download(timeout=timeout_ms) as dl_info:
                    download_btn.click()
            except PlaywrightTimeoutError as exc:
                raise MandalaFetchError("Mandala nie zwrocila pliku do pobrania w oczekiwanym czasie") from exc

            download = dl_info.value
            save_path = os.path.join(download_dir, download.suggested_filename)
            download.save_as(save_path)
            log(f"Mandala: pobrano {save_path}")

            return MandalaFetchResult(file_path=save_path, matched_hotel_label=chosen_text, warning=warning)
        finally:
            browser.close()
