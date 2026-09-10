#!/usr/bin/env python3
"""
Scraper cennika triverna.pl - ceny per pokoj, per data przyjazdu.

Uzywa oficjalnego (publicznego) GraphQL API triverna.pl (hub.triverna.pl),
z ktorego korzysta sama strona. Dla kazdej daty przyjazdu w podanym zakresie
skrypt wywoluje osobne zapytanie "calculation", ktore zwraca cene KAZDEGO
typu pokoju dostepnego w danej ofercie (nie tylko najnizsza cene widoczna
w kalendarzu na stronie).

Skrypt wyciaga tez z kalendarza oferty:
- "quantity" - jedyna informacja o dostepnosci ("liczba dostepnych pokoi
  danego dnia"), jaka triverna.pl udostepnia publicznie. Jest to wartosc
  zbiorcza dla calej oferty (API nie udostepnia rozbicia dostepnosci per
  typ pokoju).
- "minimumStay" (MLOS) - minimalna dlugosc pobytu wymagana dla danej daty
  przyjazdu.

Cena za "konkretna noc" dla pokoi INNYCH niz najtanszy jest wyliczana
przez DOKLADNE rozwiazanie (metoda najmniejszych kwadratow) ukladu rownan
zbudowanego z WSZYSTKICH nakladajacych sie okien pobytu (MLOS-owych),
jakie sprawdzilismy - patrz `_solve_overlapping_windows`. Jest to
potrzebne, bo `calculation` zwraca tylko sume calego pobytu (np. 2-3
noce), a poszczegolne noce w takim oknie moga miec rozne stawki. Takie
dokladne rozwiazanie (bez dodatkowych zapytan do API, bo dane juz mamy z
normalnego scrapowania) gwarantuje, ze suma dla kazdej faktycznie
sprawdzonej dlugosci pobytu dokladnie odtwarza prawdziwa cene z tego
okna (zweryfikowane na realnych danych: z bledu do 45 zl przy prostym
usrednianiu sasiednich okien do bledu ~0 zl).

Ponadto: rozne typy pokoi w tej samej ofercie moga miec RUZNE naturalne
oblozenie (np. "Pokoj 3-os." z defaultAdults=3) - kazdy typ pokoju jest
wyceniany przy WLASNYM, wlasciwym dla niego oblozeniu (patrz
`_room_occupancy_groups`), a nie jedna, globalna liczba osob wybrana
przez uzytkownika dla calej oferty.

Modul jest uzywany zarowno przez CLI ponizej, jak i przez triverna_gui.py.

Przyklad uzycia (CLI):

    python triverna_scraper.py \
        --url https://triverna.pl/hotel/pinea-resort-pobierowo-pobierowo \
        --start-date 2026-08-01 \
        --end-date 2026-08-31 \
        --adults 2

Domyslnie plik wyjsciowy nazywany jest automatycznie:
    <nazwa_hotelu><data_uruchomienia_dd-mm-rrrr>.csv
Mozna to nadpisac parametrem --output.

Domyslnie dlugosc pobytu dla kazdej daty przyjazdu jest brana z minimalnej
dlugosci pobytu zwracanej przez kalendarz danej oferty (MLOS). Mozna to
nadpisac parametrem --nights (wtedy dla kazdej daty uzywana jest ta sama,
stala liczba nocy).
"""

import argparse
import csv
import json
import random
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from urllib.parse import urlparse

import numpy as np
import requests

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except AttributeError:
    pass

GRAPHQL_URL = "https://hub.triverna.pl/graphql/v1.0.0"

HEADERS = {
    "Content-Type": "application/json",
    "User-Agent": "Mozilla/5.0 (compatible; TrivernaPriceScraper/1.0)",
    "Accept-Language": "pl-PL,pl;q=0.9",
}

OFFER_QUERY = """
query Offer($path: String!) {
  offer(path: $path) {
    id
    token
    active
    endDate
    mealCode
    url
    rooms {
      id
      translate { name }
      maxOccupancy
      defaultAdults
      defaultKids
    }
    hotel {
      id
      translate { name }
    }
    moreOffers {
      id
      url
      mealCode
      active
    }
  }
}
"""

CALENDAR_QUERY = """
query CalendarMonth($token: String!, $input: CalendarInput!) {
  calendar(token: $token, input: $input) {
    minimumArrivalDate
    maximumArrivalDate
    dates {
      date
      available
      noArrival
      discountedPrice
      minimumStay
      maximumStay
      quantity
    }
  }
}
"""

CALCULATION_QUERY = """
query Calc($token: String!, $input: CalculationInput!) {
  calculation(offerToken: $token, input: $input) {
    rooms {
      minPrice
      errors
      room {
        id
        translate { name }
      }
      expectedReservation {
        totalAmount
        roomPrice
        adbedPrice
        addonPrice
        touristTax
        dateFrom
        dateTo
        adults
        kids
        babies
        rooms
      }
    }
  }
}
"""


@dataclass
class CalendarDay:
    date: str
    available: bool
    no_arrival: bool
    minimum_stay: int
    maximum_stay: int
    quantity: int
    lowest_calendar_price: float


class ScrapeCancelled(Exception):
    """Podnoszone gdy uzytkownik przerwie dzialanie skryptu (np. z GUI)."""


class _RateLimiter:
    """Wymusza minimalny odstep miedzy KOLEJNYMI wyslaniami zapytan HTTP,
    wspoldzielony miedzy wszystkimi watkami roboczymi Przebiegu 1 (patrz
    max_workers w scrape()/_scrape_single_offer) - dzieki temu podniesienie
    max_workers zwieksza rownolegle "w locie" polaczenia (a wiec skraca
    czas scrapowania), ale NIE pozwala pojedynczemu watkowi wyslac zapytania
    czesciej niz raz na `delay` sekund wzgledem ostatniego wyslania - to
    jest ta sama gwarancja co dawny time.sleep(delay), tylko wspolna dla
    calej puli watkow zamiast osobna per watek.

    UWAGA: wait() zwraca PRZED faktycznym wyslaniem zapytania (pauza jest
    przed dispatchem, nie po odpowiedzi jak w starym kodzie) - to celowe,
    zeby oczekiwanie na odpowiedz serwera (roundtrip) mogl w tym czasie
    nakladac sie na dispatch innego watku. Efekt uboczny: nawet przy
    max_workers=1 tempo wysylania jest odrobine wyzsze niz w starym
    kodzie (znika "martwy" czas oczekiwania na odpowiedz z przerwy miedzy
    zapytaniami) - to swiadomy kompromis, nie usterka; jesli to budzi
    watpliwosci co do tolerancji API na wieksze obciazenie, testuj
    stopniowo zaczynajac od maleg zakresu dat."""

    def __init__(self, delay: float):
        self._delay = delay
        self._lock = threading.Lock()
        self._next_allowed = 0.0

    def wait(self) -> None:
        if self._delay <= 0:
            return
        with self._lock:
            now = time.monotonic()
            if now < self._next_allowed:
                time.sleep(self._next_allowed - now)
                now = self._next_allowed
            self._next_allowed = now + self._delay


def _parse_retry_after(value: str | None) -> float | None:
    """Parsuje naglowek Retry-After (tylko postac "liczba sekund" - format
    HTTP-date pomijamy, bo triverna.pl go nie uzywa; nieznany format ->
    None, wtedy graphql_request spada z powrotem na zwykly backoff)."""
    if not value:
        return None
    try:
        return max(0.0, float(value))
    except ValueError:
        return None


def graphql_request(
    session: requests.Session,
    query: str,
    variables: dict,
    rate_limiter: _RateLimiter,
    retries: int = 3,
    backoff: float = 1.5,
) -> dict:
    last_exc = None
    for attempt in range(1, retries + 1):
        rate_limiter.wait()
        try:
            resp = session.post(
                GRAPHQL_URL,
                headers=HEADERS,
                data=json.dumps({"query": query, "variables": variables}),
                timeout=20,
            )
            if resp.status_code == 429:
                # Osobna sciezka dla "za duzo zapytan": honorujemy Retry-After
                # gdy serwer go zwrocil, inaczej exponencjalny backoff+jitter -
                # NIE ten sam, krotki liniowy backoff co dla zwyklego 5xx/timeoutu
                # ponizej, bo to inny rodzaj bledu (przeciazenie/throttling, nie
                # chwilowa usterka pojedynczego zapytania).
                retry_after = _parse_retry_after(resp.headers.get("Retry-After"))
                wait_s = retry_after if retry_after is not None else backoff * (2**attempt) + random.uniform(0, 0.5)
                last_exc = requests.exceptions.HTTPError(
                    f"429 Too Many Requests (proba {attempt}/{retries})", response=resp
                )
                if attempt < retries:
                    time.sleep(wait_s)
                continue
            resp.raise_for_status()
            return resp.json()
        except (requests.RequestException, json.JSONDecodeError) as exc:
            last_exc = exc
            if attempt < retries:
                time.sleep(backoff * attempt)
    raise RuntimeError(f"GraphQL request failed after {retries} attempts: {last_exc}")


def hotel_path_from_url(url: str) -> str:
    """Zamienia pelny URL hotelu na sciezke ('/hotel/slug[?offer=..&category=..]')
    wymagana przez API. Zachowuje query string (offer=/category=), bo API
    honoruje go do wybrania KONKRETNEJ oferty/pakietu pod tym hotelem
    (rozne plany wyzywienia itp.) - bez tego zawsze trafialibysmy w
    domyslna oferte hotelu, ignorujac to, ktory dokladnie link wkleil
    uzytkownik."""
    parsed = urlparse(url if "://" in url else f"https://{url}")
    path = parsed.path.rstrip("/")
    if not path:
        raise ValueError(f"Nie udalo sie wyciagnac sciezki hotelu z URL: {url}")
    if parsed.query:
        path = f"{path}?{parsed.query}"
    return path


def sanitize_filename_part(name: str) -> str:
    """Usuwa znaki niedozwolone w nazwach plikow na Windows i normalizuje spacje."""
    name = name.strip()
    name = re.sub(r"\s+", "_", name)
    name = re.sub(r'[<>:"/\\|?*]', "", name)
    return name


def fetch_offer(session: requests.Session, path: str, rate_limiter: _RateLimiter) -> dict:
    data = graphql_request(session, OFFER_QUERY, {"path": path}, rate_limiter)
    if "error" in data:
        raise RuntimeError(f"Blad pobierania oferty dla {path}: {data['error']}")
    offer = data.get("data", {}).get("offer")
    if not offer:
        raise RuntimeError(f"Nie znaleziono oferty pod sciezka {path}: {data}")
    return offer


def discover_offer_graph(
    session: requests.Session, base_offer: dict, base_path: str, rate_limiter: _RateLimiter, log=print
) -> list:
    """Znajduje WSZYSTKIE aktywne oferty powiazane z ta startowa poprzez pole
    'moreOffers' (widoczne na stronie jako sekcja "Zobacz tez:") - np. ten
    sam hotel sprzedawany osobno w planie BB i HB, albo pod kilkoma roznymi
    pakietami/kategoriami. To realne, oddzielne obiekty Offer (wlasny
    token, czesto tez czesciowo inny zestaw pokoi), wiec kazda wymaga
    WLASNEGO, niezaleznego scrapowania kalendarza/cen - nie da sie ich
    wywnioskowac z jednej oferty.

    Przeszukuje graf w szerz (BFS) - powiazania bywaja dwukierunkowe/w
    pelni polaczone (kazda oferta odsyla do pozostalych), wiec ograniczamy
    sie do juz odwiedzonych ID, zeby nie petlic sie w nieskonczonosc.

    Zwraca liste (path, offer_dict) - zawsze zawiera oferte startowa jako
    pierwszy element, potem pozostale w kolejnosci odkrycia."""
    visited: dict[int, tuple] = {}
    if base_offer.get("id") is not None:
        visited[base_offer["id"]] = (base_path, base_offer)

    queue = list(base_offer.get("moreOffers") or [])
    seen_ids = set(visited.keys())
    while queue:
        entry = queue.pop(0)
        oid = entry.get("id")
        if oid is None or oid in seen_ids:
            continue
        seen_ids.add(oid)
        if not entry.get("active", True):
            continue
        entry_path = entry.get("url")
        if not entry_path:
            continue
        try:
            full_offer = fetch_offer(session, entry_path, rate_limiter)
        except Exception as exc:
            log(f"  [ostrzezenie] nie udalo sie pobrac powiazanej oferty {entry_path}: {exc}")
            continue
        if full_offer.get("id") is not None and full_offer.get("active", True):
            visited[full_offer["id"]] = (entry_path, full_offer)
        for m in full_offer.get("moreOffers") or []:
            if m.get("id") not in seen_ids:
                queue.append(m)

    return list(visited.values())


def month_range(start: date, end: date):
    """Generuje kolejne pary (rok, miesiac) pokrywajace zakres [start, end]."""
    y, m = start.year, start.month
    while (y, m) <= (end.year, end.month):
        yield y, m
        m += 1
        if m > 12:
            m = 1
            y += 1


def fetch_calendar_days(
    session: requests.Session, token: str, start: date, end: date, rate_limiter: _RateLimiter, log=print
) -> dict:
    """Pobiera dane kalendarza (dostepnosc, MLOS, ilosc wolnych pokoi) dla zakresu dat."""
    days: dict[str, CalendarDay] = {}
    for year, month in month_range(start, end):
        data = graphql_request(
            session,
            CALENDAR_QUERY,
            {"token": token, "input": {"year": year, "month": month, "startDateInterval": 0}},
            rate_limiter,
        )
        if "error" in data:
            log(f"  [ostrzezenie] kalendarz {year}-{month:02d}: {data['error']}")
            continue
        calendar = data.get("data", {}).get("calendar")
        if not calendar:
            continue
        for d in calendar["dates"]:
            days[d["date"]] = CalendarDay(
                date=d["date"],
                available=d["available"],
                no_arrival=d["noArrival"],
                minimum_stay=d["minimumStay"] or 1,
                maximum_stay=d["maximumStay"] or 99,
                quantity=d["quantity"],
                lowest_calendar_price=d["discountedPrice"],
            )
    return days


def daterange(start: date, end: date):
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def fetch_calculation(
    session: requests.Session,
    token: str,
    arrival: date,
    nights: int,
    adults: int,
    children: int,
    babies: int,
    rooms_booked: int,
    rate_limiter: _RateLimiter,
) -> dict | None:
    """Wywoluje zapytanie 'calculation' dla danej daty przyjazdu i liczby nocy.

    Zwraca RAW dane (bez zadnego przeliczania na "per noc"): slownik
    room_id -> {"room_name", "total", "tourist_tax", "errors"}, albo None
    jesli API zwrocilo blad / brak danych dla tej kombinacji.

    UWAGA: to zapytanie zwraca tylko CALKOWITA cene calego pobytu (np. za
    3 noce), NIE rozbicie na poszczegolne noce - a poszczegolne noce w
    obrebie takiego pobytu czesto MAJA ROZNE stawki (np. weekend drozszy
    niz dzien powszedni). Przeliczenie na sensowna cene "za te konkretna
    noc" (uwzgledniajace kalendarz oferty i nakladajace sie okna dat)
    dzieje sie POZNIEJ, w scrape() - patrz tam komentarz przy
    `_compute_smoothed_price_per_night`.
    """
    checkout = arrival + timedelta(days=nights)
    variables = {
        "token": token,
        "input": {
            "startDate": arrival.strftime("%Y-%m-%d"),
            "endDate": checkout.strftime("%Y-%m-%d"),
            "rooms": rooms_booked,
            "adults": adults,
            "children": children,
            "babies": babies,
            "adBedAdults": 0,
            "adBedKids": 0,
        },
    }

    data = graphql_request(session, CALCULATION_QUERY, variables, rate_limiter)

    if "error" in data:
        return None

    calc = data.get("data", {}).get("calculation")
    if not calc or not calc.get("rooms"):
        return None

    per_room = {}
    for room_entry in calc["rooms"]:
        room = room_entry.get("room") or {}
        room_id = room.get("id")
        room_name = (room.get("translate") or {}).get("name")
        reservation = room_entry.get("expectedReservation")
        errors = room_entry.get("errors") or []

        total_price = reservation["totalAmount"] if reservation else None
        tourist_tax = reservation["touristTax"] if reservation else None

        per_room[room_id] = {
            "room_name": room_name,
            "total": total_price,
            "tourist_tax": tourist_tax,
            "errors": "; ".join(str(e) for e in errors) if errors else "",
        }

    return per_room


# Grupa oblozenia (dorosli, dzieci), ktorej cena z kalendarza oferty
# (discountedPrice) zdaje sie zawsze dotyczyc - potwierdzone empirycznie
# na wielu hotelach (Pinea, Jurata, Skal, Mercure Szczyrk, Wierchomla).
REFERENCE_OCCUPANCY = (2, 0)


def _room_occupancy_groups(rooms: list, fallback_adults: int, fallback_children: int) -> dict:
    """Zwraca {room_id: (adults, kids)} - naturalne/domyslne oblozenie
    KAZDEGO pokoju wg samej oferty (pole defaultAdults/defaultKids), a nie
    jedna, globalna liczba osob wybrana przez uzytkownika.

    To wazne, bo rozne typy pokoi maja rozna naturalna pojemnosc (np.
    "Pokoj 3-os." ma defaultAdults=3) - odpytanie calculation() jedna,
    wspolna liczba doroslych dla WSZYSTKICH pokoi jednoczesnie dawalo:
    - dla pokoi z defaultAdults rownym zadanej liczbie: poprawna cene,
    - dla pokoi z INNYM defaultAdults: cene dla niewlasciwego oblozenia
      (albo, gdy zadana liczba przekraczala pojemnosc pozostalych pokoi,
      te pokoje w ogole znikaly z odpowiedzi API), przez co zapisywana
      cena nie zgadzala sie z tym, co realnie pokazuje strona.
    """
    groups = {}
    for r in rooms:
        adults = r.get("defaultAdults")
        kids = r.get("defaultKids")
        if not adults:
            adults, kids = fallback_adults, fallback_children
        groups[r["id"]] = (adults, kids or 0)
    return groups


def _pick_reference_occupancy(groups_present: set) -> tuple:
    """Wybiera, ktora z wykrytych grup oblozenia odpowiada temu, co
    zaklada kalendarz oferty (patrz REFERENCE_OCCUPANCY). Jesli zaden
    pokoj nie ma dokladnie tej konfiguracji, bierzemy oblozenie o
    najmniejszej lacznej liczbie osob jako najlepsze przyblizenie."""
    if REFERENCE_OCCUPANCY in groups_present:
        return REFERENCE_OCCUPANCY
    return min(groups_present, key=lambda g: (g[0] + g[1], g))


def _solve_overlapping_windows(equations: list) -> dict:
    """Dokladnie rozwiazuje (metoda najmniejszych kwadratow / rozwiazanie o
    minimalnej normie dla ukladu niedookreslonego) uklad rownan postaci
    "suma wartosci na tych dniach = suma_calego_okna", zbudowany z WIELU
    nakladajacych sie okien pobytu (patrz modul docstring).

    `equations`: lista (lista_dat_w_oknie, suma_dla_tego_okna).

    W przeciwienstwie do prostego usredniania lokalnych szacunkow z
    poprzedniej wersji (usrednienie 1-2 sasiednich okien), to rozwiazuje
    WSZYSTKIE dostepne rownania NARAZ, co GWARANTUJE, ze SUMA dla kazdego
    faktycznie zaobserwowanego okna (czyli kazdej mozliwej dlugosci
    pobytu >= MLOS, jaka realnie sprawdzilismy) dokladnie odtwarza
    prawdziwa cene z tego okna - zweryfikowane na realnych danych (Hotel
    Wierchomla, pokoj 3-os.): z bledu do 45 zl przy prostym usrednianiu do
    bledu ~0 zl. Pojedyncza noc pozostaje teoretycznie niedookreslona o
    jedna stala (API nie udostepnia rozbicia na pojedyncze noce ponizej
    MLOS), ale to bez znaczenia - liczy sie tylko, zeby SUMY dla realnych,
    mozliwych dlugosci pobytu byly dokladne, a te NIE zaleza od wyboru tej
    stalej (rozwiazanie o minimalnej normie po prostu dzieli ja w miare
    rownomiernie miedzy noce, dajac tez rozsadnie wygladajace pojedyncze
    wartosci).
    """
    if not equations:
        return {}
    dates_involved = sorted({d for span, _ in equations for d in span})
    idx = {d: i for i, d in enumerate(dates_involved)}
    a = np.zeros((len(equations), len(dates_involved)))
    b = np.zeros(len(equations))
    for row, (span, rhs) in enumerate(equations):
        for d in span:
            a[row, idx[d]] += 1
        b[row] = rhs
    solution, *_ = np.linalg.lstsq(a, b, rcond=None)
    return {d: float(solution[idx[d]]) for d in dates_involved}


def _window_span(day: date, nights: int) -> list:
    return [day + timedelta(days=i) for i in range(nights)]


def _iter_all_windows(window_data: dict):
    """Zwraca (data, liczba_nocy, dane_grup) dla KAZDEGO zaobserwowanego
    okna - zarowno "glownego" (tego, ktore trafia tez do wynikowych
    wierszy), jak i dodatkowej "probki" o innej dlugosci pobytu (patrz
    scrape() - `extra`), dodawanej specjalnie po to, zeby zlikwidowac
    niedookreslonosc pojedynczej doby opisana w _solve_overlapping_windows.
    Rozwiazywanie ukladu rownan (ponizej) korzysta z OBU rodzajow okien
    naraz - nie ma znaczenia, ktore jest "glowne"."""
    for d, win in window_data.items():
        yield d, win["nights"], win["groups"]
        extra = win.get("extra")
        if extra is not None:
            yield d, extra["nights"], extra["groups"]


def _solve_group_base_prices(
    window_data: dict, group: tuple, is_reference_group: bool, calendar_days: dict
) -> dict:
    """Cena bazowa "za ta noc" dla najtanszego pokoju W DANEJ GRUPIE
    oblozenia, dla kazdej daty.

    Dla grupy referencyjnej (patrz _pick_reference_occupancy) to dokladna
    cena z kalendarza oferty - sprawdzone, ze pokrywa sie co do zlotowki z
    cena na stronie. Kalendarz NIE dotyczy jednak innych oblozen (np.
    pokoju 3-os. z defaultAdults=3) - dla nich rozwiazujemy dokladnie
    (patrz _solve_overlapping_windows) cene najtanszego pokoju W TEJ
    GRUPIE wzgledem niego samego (bez zewnetrznego punktu odniesienia).
    """
    if is_reference_group:
        return {
            d: info.lowest_calendar_price
            for d_str, info in calendar_days.items()
            if (d := datetime.strptime(d_str, "%Y-%m-%d").date()) and info.lowest_calendar_price
        }

    equations = []
    for d, nights, groups_data in _iter_all_windows(window_data):
        g = groups_data.get(group)
        if g is None:
            continue
        equations.append((_window_span(d, nights), g["min_total"]))
    return _solve_overlapping_windows(equations)


def _solve_room_markups(window_data: dict, group: tuple, room_id, group_base: dict) -> dict:
    """Nadwyzka (markup) danego pokoju wzgledem ceny bazowej grupy
    (`group_base`, patrz _solve_group_base_prices), dla kazdej daty - patrz
    _solve_overlapping_windows (dokladne rozwiazanie, nie usrednianie).

    UWAGA: nadwyzke liczymy wzgledem JUZ WYLICZONEJ, spojnej sekwencji
    `group_base` (jedna wartosc na noc dla calej grupy), a NIE wzgledem
    "min_total" danego okna wprost. Powod: "najtanszy pokoj w oknie" moze
    byc RUZNYM fizycznym pokojem w roznych, nakladajacych sie oknach (np.
    tanszy pokoj bywa wyprzedany akurat na te konkretna date przyjazdu, a
    dostepny na sasiednia) - uzycie surowego min_total jako punktu
    odniesienia psulo wtedy spojnosc ukladu rownan i dawalo bledy rzedu
    kilkudziesieciu zlotych (sprawdzone na Hotel Royal Baltic: 1398 zl
    realnie vs 1338 zl w raporcie). Odniesienie do jednej, juz spojnej
    sekwencji `group_base` to naprawia.
    """
    equations = []
    for d, nights, groups_data in _iter_all_windows(window_data):
        g = groups_data.get(group)
        if g is None:
            continue
        e = g["per_room"].get(room_id)
        if e is None or e["total"] is None:
            continue
        span = _window_span(d, nights)
        if not all(s in group_base for s in span):
            continue
        base_sum = sum(group_base[s] for s in span)
        equations.append((span, e["total"] - base_sum))
    return _solve_overlapping_windows(equations)


CSV_FIELDNAMES = [
    "hotel", "offer_path", "arrival_date", "departure_date", "nights",
    "minimum_stay_nights", "room_id", "room_name", "meal_code", "adults", "children",
    "babies", "rooms_booked", "total_price", "price_per_night", "calendar_lowest_price",
    "tourist_tax", "available_quantity", "errors",
]


def _fetch_groups_data(
    session: requests.Session,
    token: str,
    day: date,
    nights: int,
    groups_present: set,
    room_occupancy: dict,
    babies: int,
    rooms_booked: int,
    rate_limiter: _RateLimiter,
    log,
    day_str: str,
    quiet: bool = False,
) -> dict:
    """Pobiera surowe dane (suma calego okna) dla KAZDEJ grupy oblozenia,
    dla jednej konkretnej (data, liczba_nocy). Wydzielone z glownej petli
    Przebiegu 1, bo scrape() wywoluje to DWA razy na date - raz dla
    "glownego" okna (MLOS), raz dla dodatkowej probki o innej dlugosci
    pobytu (patrz duzy komentarz w _scrape_single_offer)."""
    groups_data = {}
    for group in groups_present:
        g_adults, g_kids = group
        try:
            per_room = fetch_calculation(
                session, token, day, nights, g_adults, g_kids, babies, rooms_booked, rate_limiter
            )
        except RuntimeError as exc:
            # Pojedyncza data moze byc trwale zepsuta po stronie API
            # Triverny (np. brak danych cenowych daleko w przyszlosci dla
            # malo obleganego hotelu - sprawdzone na realnym przypadku:
            # Grand Hotel Tiffi zwracal blad 500 dla konkretnych dat ~1.5
            # miesiaca naprzod). Pomijamy TYLKO te date/grupe/dlugosc
            # pobytu, zamiast przerywac cale skanowanie hotelu.
            if not quiet:
                log(f"  {day_str} ({g_adults} dor.+{g_kids} dz.): blad zapytania API, pomijam te date - {exc}")
            continue
        if not per_room:
            continue
        # Zachowujemy tylko wpisy pokoi, ktorych WLASNE (naturalne)
        # oblozenie to ta grupa - reszta odpowiedzi (inne pokoje przy tym
        # samym zapytaniu) dotyczy oblozenia dla nich niewlasciwego.
        per_room = {rid: e for rid, e in per_room.items() if room_occupancy.get(rid) == group and e["total"] is not None}
        if not per_room:
            continue
        groups_data[group] = {
            "per_room": per_room,
            "min_total": min(e["total"] for e in per_room.values()),
        }
    return groups_data


def _fetch_day_data(
    session: requests.Session,
    token: str,
    day: date,
    idx: int,
    total_days: int,
    calendar_days: dict,
    groups_present: set,
    room_occupancy: dict,
    babies: int,
    rooms_booked: int,
    fixed_nights: int | None,
    rate_limiter: _RateLimiter,
    check_cancel,
) -> tuple[date, dict | None, list, bool]:
    """Jednostka pracy Przebiegu 1 dla JEDNEJ daty przyjazdu - wydzielona z
    _scrape_single_offer tak, zeby dalo sie ja uruchamiac rownolegle w puli
    watkow (patrz ThreadPoolExecutor nizej). Celowo NIE pisze nigdzie do
    dzielonego stanu (window_data/log/progress_callback) - caly wynik
    (wpis do window_data + zebrane linie logu) zwraca do zlozenia przez
    watek glowny, zeby te zapisy zawsze dzialy sie w jednym miejscu, bez
    blokad i bez ryzyka rozjechanego/cofajacego sie paska postepu.

    Nigdy nie podnosi wyjatku poza ScrapeCancelled (przerwanie przez
    uzytkownika, musi przerwac CALE skanowanie - patrz _scrape_single_offer)
    - kazdy inny blad jest lapany i zwracany jako (dzien pominiety,
    had_error=True), zeby jeden zepsuty dzien nie ubijal calego zakresu."""
    check_cancel()
    day_str = day.strftime("%Y-%m-%d")
    lines: list[str] = []
    info = calendar_days.get(day_str)

    if info is None:
        lines.append(f"[{idx}/{total_days}] {day_str}: brak danych kalendarza, pomijam")
        return day, None, lines, False
    if not info.available or info.no_arrival:
        lines.append(f"[{idx}/{total_days}] {day_str}: niedostepna jako data przyjazdu, pomijam")
        return day, None, lines, False

    nights = fixed_nights if fixed_nights else max(info.minimum_stay, 1)
    checkout = day + timedelta(days=nights)
    checkout_str = checkout.strftime("%Y-%m-%d")

    lines.append(
        f"[{idx}/{total_days}] Sprawdzam {day_str} -> {checkout_str} ({nights} noc/y, MLOS={info.minimum_stay})..."
    )

    def buffer_log(msg):
        lines.append(str(msg))

    try:
        groups_data = _fetch_groups_data(
            session, token, day, nights, groups_present, room_occupancy,
            babies, rooms_booked, rate_limiter, buffer_log, day_str,
        )

        if not groups_data:
            lines.append(f"  {day_str}: brak danych o cenie, pomijam")
            return day, None, lines, False

        # Dodatkowa "probka" o INNEJ liczbie nocy dla TEJ SAMEJ daty
        # przyjazdu - patrz duzy komentarz w oryginalnej wersji tej petli
        # (historia gita) / docstring _solve_overlapping_windows: eliminuje
        # niedookreslonosc pojedynczej doby w ukladzie rownan Przebiegu 2.
        extra_nights = nights + 1
        if info.maximum_stay and extra_nights > info.maximum_stay:
            extra_nights = nights - 1 if nights > 1 else 0
        extra_groups_data = None
        if extra_nights:
            extra_groups_data = _fetch_groups_data(
                session, token, day, extra_nights, groups_present, room_occupancy,
                babies, rooms_booked, rate_limiter, buffer_log, day_str, quiet=True,
            )
            if not extra_groups_data:
                extra_groups_data = None

        entry = {
            "nights": nights,
            "groups": groups_data,
            "checkout_str": checkout_str,
            "extra": {"nights": extra_nights, "groups": extra_groups_data} if extra_groups_data else None,
        }
        total_rooms_found = sum(len(g["per_room"]) for g in groups_data.values())
        lines.append(f"  {day_str}: OK ({total_rooms_found} pokoi)")
        return day, entry, lines, False
    except ScrapeCancelled:
        raise
    except Exception as exc:
        lines.append(f"  {day_str}: nieoczekiwany blad, pomijam ten dzien - {exc}")
        return day, None, lines, True


def _scrape_single_offer(
    session: requests.Session,
    offer: dict,
    path: str,
    start_date: date,
    end_date: date,
    adults: int,
    children: int,
    babies: int,
    rooms_booked: int,
    fixed_nights: int | None,
    rate_limiter: _RateLimiter,
    log,
    check_cancel,
    progress_callback=None,
    progress_offset: int = 0,
    progress_total: int | None = None,
    max_workers: int = 1,
) -> tuple[list[dict], str, dict]:
    """Scrapuje JEDNA konkretna oferte (patrz scrape() nizej - dla hoteli z
    kilkoma powiazanymi ofertami/pakietami to jest wywolywane osobno dla
    kazdej z nich). Zwraca (lista_wierszy, nazwa_hotelu, kalendarz_dni)."""
    token = offer["token"]
    hotel_name = offer.get("hotel", {}).get("translate", {}).get("name", path)
    meal_code = offer.get("mealCode") or ""
    room_names = {r["id"]: r["translate"]["name"] for r in offer["rooms"]}
    log(f"Hotel: {hotel_name} | typy pokoi: {', '.join(room_names.values())}")

    room_occupancy = _room_occupancy_groups(offer["rooms"], adults, children)
    groups_present = set(room_occupancy.values())
    reference_group = _pick_reference_occupancy(groups_present)
    if len(groups_present) > 1:
        groups_desc = ", ".join(
            f"{a} dor.+{k} dz.{' [referencyjna]' if (a, k) == reference_group else ''}"
            for a, k in sorted(groups_present)
        )
        log(
            f"Wykryto {len(groups_present)} rozne naturalne oblozenia pokoi ({groups_desc}) - "
            f"kazdy typ pokoju bedzie wyceniany przy WLASNYM, wlasciwym dla niego oblozeniu."
        )

    check_cancel()

    log(f"Pobieranie kalendarza dostepnosci {start_date} .. {end_date}")
    calendar_days = fetch_calendar_days(session, token, start_date, end_date, rate_limiter, log=log)

    # --- Przebieg 1: pobierz surowe dane (suma calego okna pobytu) dla kazdej daty, ---
    # osobno per grupa oblozenia (patrz _room_occupancy_groups). Rownolegle w
    # puli watkow (patrz max_workers) - to praca siecowa (I/O), wiec GIL nie
    # przeszkadza, a wspolny _RateLimiter (patrz wyzej) i tak ogranicza
    # laczne tempo wysylanych zapytan. KAZDY dzien to osobne zadanie
    # zwracajace swoj wynik (patrz _fetch_day_data) - zapis do window_data,
    # log() i progress_callback() dzieja sie WYLACZNIE na watku glownym
    # ponizej (nigdy w watku roboczym), zeby uniknac wyscigow oraz
    # rozjezdzajacego/cofajacego sie paska postepu.
    window_data: dict[date, dict] = {}
    days = list(daterange(start_date, end_date))
    total_days = len(days)
    completed = 0
    consecutive_errors = 0
    MAX_CONSECUTIVE_ERRORS = 8  # zabezpieczenie przed cichym dokonczeniem scrapowania w trakcie blokady/awarii API

    with ThreadPoolExecutor(max_workers=max(1, max_workers)) as executor:
        futures = {
            executor.submit(
                _fetch_day_data, session, token, day, idx, total_days, calendar_days,
                groups_present, room_occupancy, babies, rooms_booked, fixed_nights,
                rate_limiter, check_cancel,
            ): day
            for idx, day in enumerate(days, start=1)
        }
        try:
            for future in as_completed(futures):
                day = futures[future]
                _, entry, log_lines, had_error = future.result()
                for line in log_lines:
                    log(line)
                if entry is not None:
                    window_data[day] = entry
                completed += 1
                if progress_callback:
                    progress_callback(progress_offset + completed, progress_total or total_days, day.strftime("%Y-%m-%d"))
                if had_error:
                    consecutive_errors += 1
                    if consecutive_errors >= MAX_CONSECUTIVE_ERRORS:
                        executor.shutdown(wait=True, cancel_futures=True)
                        raise RuntimeError(
                            f"Przerwano: {consecutive_errors} kolejnych dni z rzedu zakonczylo sie "
                            "nieoczekiwanym bledem zapytania - mozliwe blokowanie/limitowanie zapytan "
                            "przez triverna.pl. Sprobuj ponownie z mniejsza liczba watkow roboczych."
                        )
                else:
                    consecutive_errors = 0
        except ScrapeCancelled:
            executor.shutdown(wait=True, cancel_futures=True)
            raise

    # --- Przebieg 2: dla kazdej grupy oblozenia i kazdego pokoju w niej, ---
    # DOKLADNIE rozwiaz (raz, dla calego zakresu dat naraz) cene bazowa i
    # nadwyzke "za noc" - patrz docstring _solve_overlapping_windows.
    # Zero dodatkowych zapytan do API - tylko na podstawie juz pobranych,
    # nakladajacych sie okien z Przebiegu 1.
    group_base_prices: dict[tuple, dict] = {}
    room_markups: dict[tuple, dict] = {}
    for group in groups_present:
        group_base_prices[group] = _solve_group_base_prices(
            window_data, group, group == reference_group, calendar_days
        )
        room_ids_in_group = set()
        for _d, _nights, groups_data in _iter_all_windows(window_data):
            g = groups_data.get(group)
            if g:
                room_ids_in_group.update(g["per_room"].keys())
        for room_id in room_ids_in_group:
            room_markups[(group, room_id)] = _solve_room_markups(
                window_data, group, room_id, group_base_prices[group]
            )

    results: list[dict] = []
    for day, win in window_data.items():
        day_str = day.strftime("%Y-%m-%d")
        info = calendar_days[day_str]
        for group, g in win["groups"].items():
            base_price = group_base_prices.get(group, {}).get(day)
            g_adults, g_kids = group
            for room_id, entry in g["per_room"].items():
                markup = room_markups.get((group, room_id), {}).get(day)
                if base_price is not None and markup is not None:
                    price_per_night = round(base_price + markup, 2)
                else:
                    price_per_night = round(entry["total"] / win["nights"], 2)
                results.append(
                    {
                        "hotel": hotel_name,
                        "offer_path": path,
                        "arrival_date": day_str,
                        "departure_date": win["checkout_str"],
                        "nights": win["nights"],
                        "minimum_stay_nights": info.minimum_stay,
                        "room_id": room_id,
                        "room_name": entry["room_name"] or room_names.get(room_id, "?"),
                        "meal_code": meal_code,
                        "adults": g_adults,
                        "children": g_kids,
                        "babies": babies,
                        "rooms_booked": rooms_booked,
                        "total_price": entry["total"],
                        "price_per_night": price_per_night,
                        "calendar_lowest_price": info.lowest_calendar_price,
                        "tourist_tax": entry["tourist_tax"],
                        "available_quantity": info.quantity,
                        "errors": entry["errors"],
                    }
                )

    log(f"Wyliczanie cen zakonczone: {len(results)} wierszy (data x pokoj) z {len(window_data)} sprawdzonych dat.")
    return results, hotel_name, calendar_days


def scrape(
    url: str,
    start_date: date,
    end_date: date,
    adults: int,
    children: int,
    babies: int,
    rooms_booked: int,
    fixed_nights: int | None,
    delay: float,
    log=print,
    progress_callback=None,
    cancel_event=None,
    max_workers: int = 1,
) -> tuple[list[dict], str, dict]:
    """Glowna funkcja scrapujaca. Zwraca (lista_wierszy, nazwa_hotelu, kalendarz_dni).

    Hotel na Trivernie moze miec KILKA powiazanych ofert/pakietow pod tym
    samym linkiem "Zobacz tez:" (np. ten sam hotel osobno w planie BB i
    HB, czasem tez pod roznymi kategoriami/promocjami) - to realnie
    ODREBNE obiekty Offer z wlasnym tokenem i czesto czesciowo innym
    zestawem pokoi, wiec nie da sie ich cen wywnioskowac z jednej
    scrapowanej oferty. Funkcja najpierw odkrywa CALY graf takich
    powiazanych ofert (patrz discover_offer_graph), a nastepnie scrapuje
    KAZDA z nich osobno i laczy wyniki - sprawdzone na realnym przypadku:
    Mercure Szczyrk Resort ma 4 takie oferty (2x BB, 2x HB).

    log: funkcja przyjmujaca jeden argument tekstowy - do wypisywania postepu.
    progress_callback: opcjonalna funkcja (aktualny_indeks, wszystkie_dni, tekst_daty).
    cancel_event: opcjonalny obiekt z metoda .is_set() (np. threading.Event) -
        gdy zwroci True, scrapowanie jest przerywane (podnoszony ScrapeCancelled).
    kalendarz_dni: slownik data_str -> CalendarDay (z PIERWSZEJ/glownej
        oferty), przydatny np. do wyliczenia MLOS oferty.
    max_workers: liczba rownoleglych watkow do pobierania cen (Przebiegu 1)
        na oferte. Domyslnie 1 (dokladnie jak w dawnej, sekwencyjnej
        wersji - jedno zapytanie na raz). Wieksze wartosci przyspieszaja
        scrapowanie, ale rownolegle otwarte polaczenia do triverna.pl to
        inny ksztalt ruchu niz dotychczasowy pojedynczy strumien - testuj
        stopniowo (male zakresy dat), zanim zwiekszysz to na stale.
    """

    def check_cancel():
        if cancel_event is not None and cancel_event.is_set():
            raise ScrapeCancelled("Przerwano przez uzytkownika")

    session = requests.Session()
    rate_limiter = _RateLimiter(delay)
    path = hotel_path_from_url(url)

    log(f"Pobieranie danych oferty: {path}")
    base_offer = fetch_offer(session, path, rate_limiter)
    check_cancel()

    offers = discover_offer_graph(session, base_offer, path, rate_limiter, log=log)
    if len(offers) > 1:
        opis = ", ".join(f"{o.get('mealCode') or '?'} ({p})" for p, o in offers)
        log(
            f"Wykryto {len(offers)} powiazanych ofert/pakietow dla tego hotelu (sekcja \"Zobacz tez\" "
            f"na stronie - np. rozne plany wyzywienia) - kazda zostanie zescrapowana oddzielnie: {opis}"
        )

    total_days_per_offer = (end_date - start_date).days + 1
    progress_total = total_days_per_offer * len(offers)

    all_results: list[dict] = []
    # Deduplikacja: gdy dwie oferty daja cene dla DOKLADNIE tego samego
    # pokoju/daty/planu wyzywienia (rzadki przypadek - np. dwie oferty BB
    # bedace w praktyce alternatywnymi pakietami tej samej stawki),
    # zachowujemy WCZESNIEJ znaleziona (oferta startowa ma pierwszenstwo),
    # zamiast pisac dwie sprzeczne ceny do tego samego miejsca.
    seen_keys = set()
    hotel_name = None
    primary_calendar_days = None
    for i, (offer_path, offer) in enumerate(offers):
        check_cancel()
        if len(offers) > 1:
            log(f"--- Oferta {i + 1}/{len(offers)}: {offer.get('mealCode') or '?'} ({offer_path}) ---")
        rows, h_name, cal_days = _scrape_single_offer(
            session, offer, offer_path, start_date, end_date, adults, children, babies,
            rooms_booked, fixed_nights, rate_limiter, log, check_cancel,
            progress_callback=progress_callback,
            progress_offset=i * total_days_per_offer,
            progress_total=progress_total,
            max_workers=max_workers,
        )
        for row in rows:
            key = (row["room_id"], row["arrival_date"], row["meal_code"])
            if key in seen_keys:
                continue
            seen_keys.add(key)
            all_results.append(row)
        if hotel_name is None:
            hotel_name = h_name
        if primary_calendar_days is None:
            primary_calendar_days = cal_days

    if len(offers) > 1:
        log(f"Laczny wynik ze wszystkich ofert: {len(all_results)} wierszy (data x pokoj x plan wyzywienia).")

    return all_results, hotel_name or path, primary_calendar_days or {}


def build_output_filename(hotel_name: str, run_date: date, suffix: str = "", ext: str = "csv") -> str:
    hotel_part = sanitize_filename_part(hotel_name) or "hotel"
    date_part = run_date.strftime("%d-%m-%Y")
    return f"{hotel_part}{date_part}{suffix}.{ext}"


def write_csv(rows: list[dict], output_path: str) -> None:
    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDNAMES, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Scraper cennika triverna.pl per data i per pokoj")
    parser.add_argument("--url", required=True, help="URL hotelu na triverna.pl, np. https://triverna.pl/hotel/pinea-resort-pobierowo-pobierowo")
    parser.add_argument("--start-date", required=True, help="Data poczatkowa (YYYY-MM-DD)")
    parser.add_argument("--end-date", required=True, help="Data koncowa (YYYY-MM-DD, wlacznie)")
    parser.add_argument("--adults", type=int, default=2, help="Liczba doroslych (domyslnie 2)")
    parser.add_argument("--children", type=int, default=0, help="Liczba dzieci")
    parser.add_argument("--babies", type=int, default=0, help="Liczba niemowlat")
    parser.add_argument("--rooms", type=int, default=1, help="Liczba rezerwowanych pokoi (domyslnie 1)")
    parser.add_argument("--nights", type=int, default=None, help="Stala liczba nocy dla kazdej daty (domyslnie: MLOS z kalendarza danej daty)")
    parser.add_argument("--delay", type=float, default=0.6, help="Minimalny odstep miedzy zapytaniami w sekundach (domyslnie 0.6)")
    parser.add_argument(
        "--max-workers", type=int, default=1,
        help="Liczba rownoleglych watkow pobierania cen (domyslnie 1 = dawne, sekwencyjne dzialanie). "
             "Wieksze wartosci przyspieszaja scrapowanie, ale zwiekszaja liczbe rownoleglych polaczen do "
             "triverna.pl - testuj stopniowo na malym zakresie dat przed uzyciem na produkcji.",
    )
    parser.add_argument("--output", default=None, help="Sciezka pliku wyjsciowego CSV (domyslnie: <nazwa_hotelu><data_dd-mm-rrrr>.csv)")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    try:
        start_date = datetime.strptime(args.start_date, "%Y-%m-%d").date()
        end_date = datetime.strptime(args.end_date, "%Y-%m-%d").date()
    except ValueError as exc:
        print(f"Niepoprawny format daty: {exc}", file=sys.stderr)
        sys.exit(1)

    if end_date < start_date:
        print("--end-date musi byc >= --start-date", file=sys.stderr)
        sys.exit(1)

    rows, hotel_name, _calendar_days = scrape(
        url=args.url,
        start_date=start_date,
        end_date=end_date,
        adults=args.adults,
        children=args.children,
        babies=args.babies,
        rooms_booked=args.rooms,
        fixed_nights=args.nights,
        delay=args.delay,
        max_workers=args.max_workers,
    )

    output_path = args.output or build_output_filename(hotel_name, date.today())
    write_csv(rows, output_path)
    print(f"\nZapisano {len(rows)} wierszy do {output_path}")


if __name__ == "__main__":
    main()
