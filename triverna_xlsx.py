#!/usr/bin/env python3
"""
Zapis wynikow triverna_scraper w formacie xlsx, na bazie dwoch szablonow:

- templates/mandala_szablon.xlsx  -> cennik "Mandala": jedna wartosc MLOS dla
  calej oferty (wpisywana w naglowku), a nizej siatka data x pokoj z cena za
  noc. Szablon ma miejsce na do 8 typow pokoi (dwa bloki po 4 kolumny).
- templates/ava_report_szablon.xlsx -> raport dostepnosci "Ava": siatka
  data x pokoj z liczba dostepnych pokoi. Triverna udostepnia publicznie
  TYLKO zbiorcza (nie per-pokoj) liczbe dostepnych pokoi dla danej daty - ta
  sama wartosc jest wiec wpisywana w kazda kolumne pokoju (patrz komentarz
  dopisywany do komorki B1 w wygenerowanym pliku).

Oba szablony nie zawieraja formul - sa to czyste "siatki" danych, wiec nie
jest wymagane przeliczanie (recalc) po zapisie.
"""

import re
import sys
from copy import copy as copy_style
from datetime import date, datetime, timedelta
from statistics import multimode

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill

# --------------------------------------------------------------------- sciezki

def resource_path(*parts: str) -> str:
    """Zwraca sciezke do zasobu, dzialajaca zarowno w trybie skryptu jak i
    w spakowanym .exe (PyInstaller onefile rozpakowuje dane do sys._MEIPASS)."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base = sys._MEIPASS
    else:
        import os

        base = os.path.dirname(os.path.abspath(__file__))
    import os

    return os.path.join(base, *parts)


MANDALA_TEMPLATE_PATH = resource_path("templates", "mandala_szablon.xlsx")
AVA_TEMPLATE_PATH = resource_path("templates", "ava_report_szablon.xlsx")

# --------------------------------------------------------------- uklad szablonow

MANDALA_SHEET = "Regular"
# W surowym szablonie kolumny pokoi byly ulozone w dwoch blokach po 4 (B-E,
# G-J) z pusta kolumna-separatorem (F) i wlasna komorka MLOS na poczatku
# KAZDEGO bloku (B1, G1) - uzytkownik chce jednak ciaglych danych bez
# przerwy, wiec ten uklad jest teraz traktowany jako JEDEN ciagly blok 8
# kolumn (B..I); dawna kolumna-separator (F) i dawny poczatek drugiego
# bloku (G) sa "ucywilizowywane" (dopasowywane stylem do reszty) przy
# zapisie - patrz _harmonize_mandala_gap_columns w write_mandala_xlsx.
MANDALA_ROOM_COLS = ["B", "C", "D", "E", "F", "G", "H", "I"]
MANDALA_FORMER_GAP_COLS = ["F", "G"]
MANDALA_STYLE_REFERENCE_COL = "D"
MANDALA_MLOS_CELL = "B1"
MANDALA_DATE_COL = "A"
MANDALA_DATA_START_ROW = 10
MANDALA_HEADER_ROWS_TO_CLEAR = [2, 3, 4, 5, 6, 7, 8]  # Room/Room ID/Board/.../Internal Name
WEEKEND_FILL_MANDALA = "FFE8F5E9"

# Etykiety naglowka w kolumnie A - kolejnosc wierszy rozni sie miedzy pustym
# szablonem a "prawdziwymi" plikami z systemu (np. Board/Board ID moze byc
# przed albo po Room/Room ID), dlatego pozycje wierszy przy scalaniu
# wykrywane sa dynamicznie, a nie zakladane na sztywno.
MANDALA_HEADER_LABELS = [
    "MLOS", "Board", "Board ID", "Room", "Room ID",
    "Package", "Package ID", "Internal Name", "date",
]

AVA_SHEET = "Sheet"
AVA_ROOM_COLS = ["C", "D"]  # blok "Current availability"
AVA_ROOM_COLS_CHANGES = ["G", "H"]  # blok "Please make changes"
AVA_DATE_COL = "B"
AVA_MIRROR_DATE_COL = "F"  # blok "Please make changes" ma wlasna, mirrorowana kolumne dat
AVA_DATA_START_ROW = 3
WEEKEND_FILL_AVA = "FFDDDDDD"


class TemplateCapacityWarning(Warning):
    pass


def compute_offer_mlos(calendar_days: dict, log=print) -> int:
    """Wyznacza jedna, reprezentatywna wartosc MLOS dla calej oferty
    (najczestsza minimalna dlugosc pobytu wsrod dostepnych dat)."""
    values = [d.minimum_stay for d in calendar_days.values() if d.available and not d.no_arrival]
    if not values:
        return 1
    modes = multimode(values)
    mlos = min(modes)
    if len(set(values)) > 1:
        log(f"  Uwaga: MLOS rozni sie w zaleznosci od daty (wartosci: {sorted(set(values))}); "
            f"do szablonu wpisano najczestsza wartosc: {mlos}")
    return mlos


def _unique_rooms_in_order(rows: list[dict]) -> list[tuple[int, str]]:
    seen: dict = {}
    for r in rows:
        rid = r["room_id"]
        if rid not in seen:
            seen[rid] = r["room_name"]
    return list(seen.items())


def _unique_room_boards_in_order(rows: list[dict]) -> list[tuple[int, str, str]]:
    """Jak _unique_rooms_in_order, ale KAZDY odrebny plan wyzywienia tego
    samego pokoju (np. BB i HB - patrz scrape()/discover_offer_graph, hotel
    moze miec kilka powiazanych ofert z naprawde roznymi cenami per plan)
    liczy sie jako ODREBNA pozycja, nie jest zwijany do jednego wpisu.
    Bez tego pokoj dostepny w kilku planach wyzywienia dostawalby TYLKO
    JEDNA kolumne (i jeden, przypadkowy - zwykle pierwszy napotkany - plan
    wyzywienia w etykiecie), tracac cala reszte cen."""
    seen: dict = {}
    for r in rows:
        key = (r["room_id"], r["meal_code"])
        if key not in seen:
            seen[key] = r["room_name"]
    return [(rid, name, meal) for (rid, meal), name in seen.items()]


def _date_row_map(ws, col_letter: str, start_row: int) -> dict:
    mapping = {}
    row = start_row
    max_row = ws.max_row
    while row <= max_row:
        val = ws[f"{col_letter}{row}"].value
        if val is None:
            break
        if isinstance(val, datetime):
            mapping[val.date()] = row
        elif isinstance(val, date):
            mapping[val] = row
        row += 1
    return mapping


def _extend_date_rows(ws, col_letters: list[str], start_row: int, target_end: date, weekend_fill_hex: str, log=print) -> dict:
    """Jesli szablon nie siega docelowej daty koncowej, dopisuje kolejne
    wiersze z datami (zachowujac formatowanie weekendow: piatek+sobota)."""
    mapping = _date_row_map(ws, col_letters[0], start_row)
    if not mapping:
        return mapping

    last_date = max(mapping)
    last_row = mapping[last_date]
    if target_end <= last_date:
        return mapping

    fill_weekend = PatternFill(start_color=weekend_fill_hex, end_color=weekend_fill_hex, fill_type="solid")

    current_date = last_date
    current_row = last_row
    log(f"  Szablon nie obejmowal daty {target_end} - dopisuje brakujace wiersze...")
    while current_date < target_end:
        current_date += timedelta(days=1)
        current_row += 1
        is_weekend = current_date.weekday() in (4, 5)  # piatek=4, sobota=5 (Mon=0)
        for col in col_letters:
            cell = ws[f"{col}{current_row}"]
            cell.value = datetime(current_date.year, current_date.month, current_date.day)
            cell.number_format = "yyyy-mm-dd"
            if is_weekend:
                cell.fill = fill_weekend
        mapping[current_date] = current_row
    return mapping


def _rebase_template_dates_from_today(ws, col_letters: list[str], start_row: int, weekend_fill_hex: str) -> None:
    """Pusty szablon (mandala_szablon.xlsx / ava_report_szablon.xlsx) ma
    zaszyte na sztywno daty od dnia jego stworzenia - bez tego kroku
    KAZDY wygenerowany plik zaczynalby sie od tej samej, coraz bardziej
    nieaktualnej daty, zamiast od dnia faktycznego wygenerowania raportu.
    Nadpisuje wiec daty JUZ obecne w szablonie (od start_row w dol) tak,
    zeby pierwszy wiersz odpowiadal DZISIEJSZEJ dacie - zachowujac liczbe
    istniejacych wierszy (tylko przesuwa cala sekwencje o tyle dni, ile
    trzeba); _extend_date_rows() nizej dopisuje dalsze wiersze, jesli
    zakres scrapowania siega dalej niz to, co szablon juz mial."""
    existing = _date_row_map(ws, col_letters[0], start_row)
    if not existing:
        return
    row_count = len(existing)
    today = date.today()
    fill_weekend = PatternFill(start_color=weekend_fill_hex, end_color=weekend_fill_hex, fill_type="solid")
    fill_none = PatternFill(fill_type=None)
    for i in range(row_count):
        row = start_row + i
        current_date = today + timedelta(days=i)
        is_weekend = current_date.weekday() in (4, 5)  # piatek=4, sobota=5 (Mon=0)
        for col in col_letters:
            cell = ws[f"{col}{row}"]
            cell.value = datetime(current_date.year, current_date.month, current_date.day)
            cell.number_format = "yyyy-mm-dd"
            cell.fill = fill_weekend if is_weekend else fill_none


# ------------------------------------------------------------------- Mandala

def write_mandala_xlsx(rows: list[dict], calendar_days: dict, output_path: str, log=print) -> None:
    """Zapisuje cennik w formacie szablonu Mandala (jedna wartosc MLOS + siatka data x pokoj)."""
    if not rows:
        raise ValueError("Brak danych do zapisania (pusta lista wierszy)")

    wb = load_workbook(MANDALA_TEMPLATE_PATH)
    ws = wb[MANDALA_SHEET]

    # Ujednolic wyglad dawnej kolumny-separatora (F) i dawnego poczatku
    # drugiego bloku (G) tak, zeby wygladaly identycznie jak zwykla
    # "wewnetrzna" kolumna danych (D) - inaczej te dwie kolumny mialyby
    # brakujace/niespojne podswietlenie naglowkow, mimo ze teraz trzymaja
    # normalne dane pokoju (patrz komentarz przy MANDALA_ROOM_COLS).
    for target_col in MANDALA_FORMER_GAP_COLS:
        for r in range(1, 9):
            src = ws[f"{MANDALA_STYLE_REFERENCE_COL}{r}"]
            dst = ws[f"{target_col}{r}"]
            dst.fill = copy_style(src.fill)
            dst.border = copy_style(src.border)
            dst.font = copy_style(src.font)
            dst.alignment = copy_style(src.alignment)
            dst.number_format = src.number_format
        # Wiersz 1 (MLOS) trzymal wczesniej "resztkowe" wartosci specyficzne
        # dla dawnego ukladu dwoch blokow (F1: sama spacja jako separator,
        # G1: placeholder "MLOS: wpisz liczbe" dla drugiego bloku) - teraz
        # jest tylko JEDNA wspolna komorka MLOS (patrz MANDALA_MLOS_CELL),
        # wiec obie musza zostac jawnie wyczyszczone.
        ws[f"{target_col}1"] = None

    # F (dawna pusta kolumna-separator) nie mialo tekstow placeholder ("DP
    # uzupelni" dla Room ID/Board ID/Package/Package ID/Internal Name) -
    # G (dawny poczatek drugiego bloku) mial je juz od poczatku, wiec
    # kopiujemy je tylko dla F, zeby wygladala identycznie jak reszta
    # uzywanych kolumn pokoi.
    for r in (3, 5, 6, 7, 8):
        ws["F" + str(r)] = ws[f"{MANDALA_STYLE_REFERENCE_COL}{r}"].value

    room_order = _unique_room_boards_in_order(rows)
    if len(room_order) > len(MANDALA_ROOM_COLS):
        log(f"  [UWAGA] oferta ma {len(room_order)} kombinacji pokoj+wyzywienie - szablon Mandala "
            f"obsluguje max. {len(MANDALA_ROOM_COLS)}. Eksportuje tylko pierwsze "
            f"{len(MANDALA_ROOM_COLS)}.")
        room_order = room_order[: len(MANDALA_ROOM_COLS)]

    mlos = compute_offer_mlos(calendar_days, log=log)

    room_col: dict = {}
    for idx, (room_id, room_name, room_meal_code) in enumerate(room_order):
        col = MANDALA_ROOM_COLS[idx]
        room_col[(room_id, room_meal_code)] = col
        ws[f"{col}2"] = room_name
        ws[f"{col}4"] = room_meal_code

    used_cols = set(room_col.values())
    for col in MANDALA_ROOM_COLS:
        if col not in used_cols:
            for r in MANDALA_HEADER_ROWS_TO_CLEAR:
                ws[f"{col}{r}"] = None

    # Kolumna J trzymala w surowym szablonie 4. kolumne dawnego drugiego
    # bloku ("Wpisz nazwe pokoju 4" itd.) - teraz, gdy dane pisane sa w
    # ciagu az do I, J zostaje calkowicie poza uzywanym zakresem i musi
    # zostac jawnie wyczyszczona, inaczej zostawilaby "osierocony" naglowek
    # placeholdera zaraz za realnymi danymi.
    for r in MANDALA_HEADER_ROWS_TO_CLEAR:
        ws[f"J{r}"] = None

    ws[MANDALA_MLOS_CELL] = mlos

    _rebase_template_dates_from_today(ws, [MANDALA_DATE_COL], MANDALA_DATA_START_ROW, WEEKEND_FILL_MANDALA)
    end_date = max(datetime.strptime(r["arrival_date"], "%Y-%m-%d").date() for r in rows)
    date_map = _extend_date_rows(ws, [MANDALA_DATE_COL], MANDALA_DATA_START_ROW, end_date, WEEKEND_FILL_MANDALA, log=log)

    written = 0
    for r in rows:
        arrival = datetime.strptime(r["arrival_date"], "%Y-%m-%d").date()
        row_num = date_map.get(arrival)
        col = room_col.get((r["room_id"], r["meal_code"]))
        if row_num is None or col is None:
            continue
        cell = ws[f"{col}{row_num}"]
        cell.value = r["price_per_night"]
        cell.number_format = "0.00"
        written += 1

    wb.save(output_path)
    log(f"Zapisano cennik (Mandala): {output_path} ({written} cen, MLOS={mlos})")


# ------------------------------------------------ scalanie z przeslanym plikiem

_DIACRITICS_MAP = str.maketrans({
    "ą": "a", "ć": "c", "ę": "e", "ł": "l", "ń": "n", "ó": "o", "ś": "s", "ź": "z", "ż": "z",
    "Ą": "a", "Ć": "c", "Ę": "e", "Ł": "l", "Ń": "n", "Ó": "o", "Ś": "s", "Ź": "z", "Ż": "z",
})

# Ograniczony slownik "rozpoznawalnych" slow kluczowych typu pokoju - to na
# nich opiera sie dopasowanie nazw pokoi miedzy Triverna (PL) a przeslanym
# plikiem (czesto EN, czasem czysto PL) - te slowa sa zazwyczaj
# identyczne/zapozyczone w obu jezykach (np. "classic", "twin", "suite")
# albo ich polskimi odpowiednikami (np. "komfort", "apartament"), w
# przeciwienstwie do reszty nazwy (rodzajniki, "with", "room", "widok"
# itp.), ktora nic nie wnosi do dopasowania.
_ROOM_KEYWORD_VOCAB = {
    "classic", "twin", "queen", "double", "family", "superior", "privilege",
    "suite", "executive", "junior", "standard", "deluxe", "comfort", "studio",
    "apartment", "triple", "single", "king", "premium", "economy",
    # polskie odpowiedniki / warianty pisowni napotkane w realnych plikach
    # (np. hotele opisujace pokoje wylacznie po polsku - bez tego wpisu
    # takie hotele nie mialy ZADNEGO dopasowania i ceny w ogole sie nie
    # aktualizowaly):
    "komfort", "apartament", "delux", "dlux", "lux", "luks", "standardowy",
    "ekonomiczny", "rodzinny", "dwupoziomowy",
    # cechy pokoju, ktore realnie odrozniaja warianty tego samego "tieru"
    # (np. "Lux" vs "Lux z aneksem kuchennym"):
    "aneksem", "aneks", "kuchennym", "kuchnia", "balkonem", "balkon",
    "taras", "tarasem",
    # modyfikatory tieru - bez nich np. "Comfort" i "Comfort Plus" tokenizuja
    # sie IDENTYCZNIE (oba do samego "comfort"), co dawalo falszywy remis i
    # w efekcie pomijalo/nadpisywalo zla cene (sprawdzone na realnym
    # przypadku: Rezydencja Gubalowka, "Comfort" vs "Comfort Plus"):
    "plus", "mini", "max", "grand", "elite", "select", "prestige",
    "gold", "silver", "platinum", "vip", "exclusive", "signature", "special",
    # dodatkowe slowa napotkane przy walidacji na >20 realnych hotelach
    # (bulk test) - bez nich cale typy pokoi w ogole sie nie dopasowywaly:
    "extra", "elegance", "smart", "domek", "duo", "sypialnia",
    "skrzydlo", "nowym", "nowy", "nowe",
    # kanoniczny rdzen (bez koncowki przypadka) dodawany przez stemming w
    # _normalize_room_tokens dla "historyczny/historyczne/historycznym/..." -
    # musi tu byc, bo inaczej "words & _ROOM_KEYWORD_VOCAB" po cichu odrzuca
    # sam token rdzenia, mimo ze stemming go poprawnie wygenerowal:
    "historyczn",
    # nazwy budynkow/skrzydel w obiektach z kilkoma budynkami (np. Five
    # Seasons One/Two/Three) - bez tego rozne budynki tego samego tieru
    # tokenizuja sie identycznie i dopasowanie jest odrzucane jako
    # niejednoznaczne. Celowo BEZ "four"/"five"/"six" - "five" tu jest
    # czescia nazwy marki hotelu ("Five Seasons") i wstrzykuje szum do
    # KAZDEJ nazwy pokoju w pliku, obnizajac marginesy miedzy tierami:
    "one", "two", "three",
}


def _normalize_room_tokens(name: str) -> set:
    name = (name or "").lower().translate(_DIACRITICS_MAP)
    # rozklad dorosli+dzieci (np. "2+1", "2 + 2") to inny rodzaj informacji
    # niz prosta liczba miejsc w pokoju ("Pokoj 5-os.") i w praniu okazalo
    # sie, ze nie da sie go bezpiecznie zakodowac jako osobny token: albo
    # koliduje z plikiem, ktory w ogole nie ma cyfr (Mercure Szczyrk Resort,
    # "Privilege family Twin" obnizone ponizej progu przez dwie luzne cyfry),
    # albo - jako odrebny token typu "2+2" - PRZESTAJE przypadkiem pokrywac
    # sie z plikowa cyfra "2" z "2-os.", z ktora do tej pory poprawnie sie
    # utozsamial (sprawdzone na realnym przypadku: Magnus Resort, "Apartament
    # superior 2+2" - jako osobny token traci margines nad blizniaczym
    # "Family Superior" i zostaje bledniej odrzucony). Dlatego rozklad
    # dorosli+dzieci jest w calosci usuwany (nie tokenizowany wcale) - to
    # jedyny wariant, ktory naprawia oba te przypadki NIE psujac przy tym
    # zadnego z pozostalych sprawdzonych hoteli (potwierdzone recznym
    # przeliczeniem Jaccarda dla kazdego z 9 problematycznych hoteli z bulk
    # testu przed wdrozeniem):
    name_wo_combo = re.sub(r"\d+\s*\+\s*\d+", " ", name)
    words = set(re.findall(r"[a-z]+", name_wo_combo))
    # liczby osob (np. "Pokoj 5-os.") - niektore hotele roznicuja pokoje
    # WYLACZNIE liczba miejsc, wiec cyfry licza sie jako tokeny tak samo
    # jak slowa ze slownika (bez tego takie nazwy nie mialy zadnych
    # rozpoznawalnych tokenow i byly w calosci pomijane):
    digits = set(re.findall(r"\d+", name_wo_combo))
    # odmiana przez przypadki slowa "historyczny" (historycznym, historycznej
    # itp.) - plik czesto uzywa innego przypadka niz mianownik z Triverny, a
    # dokladne dopasowanie tekstowe do slownika by to przegapilo (sprawdzone
    # na realnym przypadku: Arche Metalowiec Muszyna). Celowo NIE odmieniamy
    # w ten sam sposob "skrzydlo" (wing) - sprawdzone, ze to tworzy falszywe
    # dopasowanie miedzy "nowym" a "historycznym" skrzydlem, bo samo "wing"
    # jest tam wspolne, a rozroznia je tylko modyfikator:
    if any(w.startswith("historyczn") for w in words):
        words.add("historyczn")
    return (words & _ROOM_KEYWORD_VOCAB) | digits


# Wiele hoteli sprzedaje TEN SAM pokoj pod kilkoma planami wyzywienia (np.
# BB i HB), wiec ich nazwa w przeslanym pliku powtarza sie w kilku
# kolumnach (roznych blokach "Board"). To daje remisy w dopasowaniu po
# samej nazwie pokoju - rozstrzygane ponizej po tekscie wyzywienia.
#
# Uwaga: "obiadokolacja" (polski termin na kolacje serwowana jak obiad)
# zawiera w sobie tez sniadanie (HB = BB + obiadokolacja), wiec samo
# sprawdzenie "czy tekst zawiera slowo kluczowe BB" daloby falszywie
# pozytywny wynik rowniez dla opisu HB. Dlatego kazda regula ma liste
# wykluczajaca ("none") - fraz, ktorych obecnosc oznacza bogatszy plan
# wyzywienia niz szukany kod.
_BOARD_CODE_RULES = {
    "RO": {"any": ["bez wyzywienia", "no meal", "room only", "self catering"], "none": []},
    "BB": {
        "any": ["sniadania", "sniadanie", "breakfast"],
        "none": ["obiadokolacj", "polpensjonat", "half board", "pelne wyzywienie", "full board", "all inclusive"],
    },
    "HB": {
        "any": ["obiadokolacj", "polpensjonat", "half board"],
        "none": ["pelne wyzywienie", "full board", "all inclusive"],
    },
    "FB": {"any": ["pelne wyzywienie", "full board"], "none": ["all inclusive"]},
    "AI": {"any": ["all inclusive"], "none": ["ultra"]},
    "UAI": {"any": ["ultra all inclusive"], "none": []},
}


def _board_matches(meal_code: str, board_text: str) -> bool:
    if not meal_code or not board_text:
        return False
    rule = _BOARD_CODE_RULES.get(meal_code.upper())
    if not rule:
        return False
    normalized = board_text.lower().translate(_DIACRITICS_MAP)
    if not any(kw in normalized for kw in rule["any"]):
        return False
    if any(kw in normalized for kw in rule["none"]):
        return False
    return True


def match_rooms(scraped_rooms: list, file_rooms: list, meal_code: str = None, log=print) -> dict:
    """Dopasowuje pokoje ze scrapowania (room_id, room_name) do kolumn w
    przeslanym pliku (col_letter, room_name, board_text_or_None).

    Zwraca dict room_id -> lista kolumn (list[str]) - zazwyczaj jedna
    kolumna, ale patrz nizej przypadek tego samego pokoju/wyzywienia
    powtorzonego w kilku blokach (np. rozne progi MLOS).

    Nazwy pokoi czesto pochodza z dwoch roznych systemow (np. Triverna po
    polsku vs. wewnetrzny plik po angielsku, albo oba po polsku, ale inaczej
    nazwane) - dokladne dopasowanie tekstowe zwykle zawiedzie. Dopasowanie
    odbywa sie wiec na podstawie wspolnych "rozpoznawalnych" slow
    kluczowych typu pokoju (patrz _ROOM_KEYWORD_VOCAB) i TYLKO gdy jest
    jednoznaczne - w przeciwnym razie pokoj jest pomijany (nie zgadujemy),
    zeby nie podmienic ceny w zlej kolumnie.

    Ten sam pokoj czasem powtarza sie w pliku w kilku kolumnach:
    - pod roznymi planami wyzywienia (np. BB i HB) - to GENUINE rozne
      produkty o roznej cenie, wiec remis rozstrzygany jest wedlug
      zgodnosci z meal_code oferty (np. "BB" -> szukamy w "Board" slowa
      "Śniadania"); jesli nie da sie rozstrzygnac, pomijamy.
    - pod tym samym planem wyzywienia, ale w innym bloku MLOS (np. "MLOS: 2"
      i "MLOS: 4") - to zazwyczaj administracyjne powielenie tego samego
      pokoju/stawki, wiec (poniewaz nasza wlasna cena per noc nie rozroznia
      dlugosci pobytu) wpisujemy ja do WSZYSTKICH takich kolumn naraz.
    """
    mapping: dict = {}

    if len(scraped_rooms) == 1 and len(file_rooms) == 1:
        (room_id, room_name), (col, file_name, _board) = scraped_rooms[0], file_rooms[0]
        mapping[room_id] = [col]
        log(f"  Dopasowanie pokoju: '{room_name}' -> kolumna {col} ('{file_name}') [jedyny pokoj po obu stronach]")
        return mapping

    file_tokens_raw = [(col, name, board, _normalize_room_tokens(name)) for col, name, board in file_rooms]

    # Slowa/cyfry wspolne dla LITERALNIE KAZDEGO odrebnego typu pokoju w tym
    # pliku (np. "apartament" i cyfra "2" u hotelu, gdzie kazdy apartament
    # jest 2-osobowy) nigdy nie moga niczego rozroznic miedzy tymi pokojami -
    # tylko zasmiecaja mianownik Jaccarda i sztucznie obnizaja pewnosc/margines
    # miedzy prawdziwie roznymi tierami (sprawdzone na realnym przypadku: Blu
    # Apartments, "Smart Suite" vs "Smart Suite Plus"). Pomijamy je przy
    # liczeniu wyniku danej pary - ale TYLKO gdy po ich usunieciu wciaz
    # zostaje jakis wspolny token miedzy TA KONKRETNA para; jesli usuniecie
    # zniszczyloby JEDYNE realne powiazanie (np. plikowa kolumna nie ma
    # zadnego innego rozpoznawalnego slowa poza tym "wspolnym" tokenem),
    # wracamy do surowych tokenow dla tej pary - inaczej tracimy dopasowanie,
    # ktore dzialalo poprawnie zanim dodano to czyszczenie (sprawdzone na
    # realnym przypadku: Hotel Orle, "Pokój 2-os. Standard" - jedyny lacznik
    # z plikowa kolumna to sama wspolna dla calego hotelu cyfra "2-osobowy").
    distinct_signatures = {tuple(sorted(tok)) for _, _, _, tok in file_tokens_raw if tok}
    common_tokens = set.intersection(*(set(sig) for sig in distinct_signatures)) if len(distinct_signatures) >= 2 else set()

    # Analogiczny szum po stronie scrapowania: gdy KAZDY pokoj w tym hotelu ma
    # ta sama (jedyna) liczbe osob bazowych (np. wszystkie pokoje sa "dla 2
    # osob", tylko rozne warianty dostawek/rozkladow), ta cyfra tez niczego nie
    # rozroznia MIEDZY pokojami tego hotelu - a ponieważ rozklad "N+M" jest
    # usuwany w calosci (patrz _normalize_room_tokens), podczas gdy zwykle
    # "N-os." zachowuje goly cyfrowy token, to samo "2" bez tej poprawki
    # niesymetrycznie zasmieca tylko CZESC pokoi, przypadkiem zrownujac wynik
    # zwyklego pokoju z jego wariantem+modyfikatorem (sprawdzone na realnym
    # przypadku: Mercure Szczyrk Resort - "Pokój 2-os. classic queen" fałszywie
    # remisowal z "Pokój 2+1 classic queen family" o ta sama kolumne, mimo ze
    # plikowa nazwa w ogole nie ma cyfr).
    scraped_digit_values = set()
    for _, name in scraped_rooms:
        scraped_digit_values |= {t for t in _normalize_room_tokens(name) if t.isdigit()}
    if len(scraped_digit_values) == 1:
        common_tokens = common_tokens | scraped_digit_values

    def _pair_score(scraped_raw: set, file_raw: set):
        s = scraped_raw - common_tokens or scraped_raw
        f = file_raw - common_tokens or file_raw
        overlap = s & f
        if overlap:
            return len(overlap) / len(s | f)
        overlap = scraped_raw & file_raw
        if not overlap:
            return None
        return len(overlap) / len(scraped_raw | file_raw)

    EPS = 1e-9

    # Faza 1: dla kazdego pokoju liczymy TYMCZASOWA propozycje dopasowania,
    # NIE wykluczajac jeszcze zadnych kolumn - kolejnosc przetwarzania pokoi
    # nie moze decydowac, kto "pierwszy" zdobedzie kolumne (patrz faza 2).
    proposals = []

    for room_id, room_name in scraped_rooms:
        scraped_tok = _normalize_room_tokens(room_name)
        if not scraped_tok:
            log(f"  [pomijam] pokoj '{room_name}' - brak rozpoznawalnych slow kluczowych do dopasowania")
            continue

        scored = []
        for col, file_name, board, tok in file_tokens_raw:
            if not tok:
                continue
            score = _pair_score(scraped_tok, tok)
            if score is None:
                continue
            scored.append((score, col, file_name, board, tok))
        scored.sort(key=lambda t: t[0], reverse=True)

        if not scored:
            log(f"  [pomijam] pokoj '{room_name}' - nie znaleziono odpowiadajacej kolumny w przeslanym pliku")
            continue

        best_score = scored[0][0]
        distinct_all_names = {c[2] for c in scored}

        # Awaryjne dopasowanie ponizej progu 0.5: jesli WSZYSTKIE rozpoznane
        # slowa kluczowe krotkiej nazwy ze scrapowania sa w calosci zawarte w
        # jedynym (zerowa konkurencja - zaden INNY, ROZNY plikowy pokoj nie
        # dzieli choc jednego tokenu) kandydacie, akceptujemy mimo niskiego
        # wyniku Jaccarda - to typowo skrotowa nazwa Triverny vs. rozbudowany
        # opis w pliku, nie prawdziwa niepewnosc (sprawdzone na realnym
        # przypadku: Hotel Orle, "Pokój DUO" - zero konkurencji, ale niski
        # wynik Jaccarda tylko przez dlugosc opisu w pliku).
        fallback_containment = (
            best_score < 0.5
            and len(distinct_all_names) == 1
            and scraped_tok <= scored[0][4]
        )

        if best_score < 0.5 and not fallback_containment:
            log(f"  [pomijam - niejednoznaczne] pokoj '{room_name}' - najlepszy kandydat '{scored[0][2]}' "
                f"(pewnosc {best_score:.2f}), zbyt niepewne dopasowanie")
            continue

        tied_cols = {c[1] for c in scored if best_score - c[0] < EPS}
        tied = [c for c in scored if c[1] in tied_cols]

        distinct_names = {c[2] for c in tied}
        distinct_boards = {c[3] for c in tied}

        if len(distinct_names) > 1:
            # Rozne (nie identyczne) nazwy pokoi dostaly przypadkiem ten
            # sam wynik dopasowania - najczesciej luka w slowniku slow
            # kluczowych (np. "Comfort" vs "Comfort Plus" bez slowa "plus"
            # w slowniku wygladaja identycznie). To PRAWDZIWA
            # niejednoznacznosc - nie wolno zgadywac ktora kolumna jest
            # wlasciwa, bo w przeciwnym razie nadpisalibysmy cene w zlym,
            # ale zajetym juz przez ten remis miejscu (sprawdzone na
            # realnym przypadku: Rezydencja Gubalowka).
            resolved = []
        elif len(distinct_boards) <= 1:
            # Identyczna nazwa pokoju i to samo wyzywienie powtorzone w
            # kilku kolumnach (np. rozne bloki MLOS) - to nie jest
            # prawdziwa niejednoznacznosc, wiec wpisujemy cene do
            # wszystkich naraz.
            resolved = tied
        elif meal_code:
            board_matched = [c for c in tied if _board_matches(meal_code, c[3])]
            resolved = board_matched if len(board_matched) >= 1 else []
        else:
            resolved = []

        if not resolved:
            if len(distinct_names) > 1:
                log(f"  [pomijam - niejednoznaczne] pokoj '{room_name}' - kilka ROZNYCH kolumn z takim samym "
                    f"wynikiem dopasowania ({', '.join(sorted(distinct_names))}) - zbyt niepewne, zeby zgadywac")
            else:
                log(f"  [pomijam - niejednoznaczne] pokoj '{room_name}' - kilka jednakowo pasujacych kolumn "
                    f"o roznym wyzywieniu ({', '.join(sorted(distinct_boards, key=lambda b: b or ''))}), "
                    f"nie mozna jednoznacznie wybrac")
            continue

        if not fallback_containment:
            remaining = [c for c in scored if c[1] not in tied_cols]
            second_score = remaining[0][0] if remaining else 0.0
            if (best_score - second_score) < 0.15:
                log(f"  [pomijam - niejednoznaczne] pokoj '{room_name}' - najlepszy kandydat '{resolved[0][2]}' "
                    f"(pewnosc {best_score:.2f}), zbyt niepewne dopasowanie")
                continue

        chosen_cols = [c[1] for c in resolved]
        chosen_names = ", ".join(dict.fromkeys(c[2] for c in resolved))
        proposals.append({
            "room_id": room_id, "room_name": room_name,
            "cols": chosen_cols, "names": chosen_names, "score": best_score,
        })

    # Faza 2: rozstrzygamy kolizje kolumn wedlug SILY dopasowania (malejaco),
    # a nie kolejnosci na liscie pokoi ze scrapowania - slabiej pasujacy
    # pokoj (np. ten sam pokoj + dodatkowy modyfikator w nazwie, jak
    # "... Twin" vs "... family Twin") nigdy nie "kradnie" kolumny nalezacej
    # realnie do mocniej pasujacego pokoju tylko dlatego, ze zostal
    # przetworzony wczesniej (sprawdzony realny przypadek: Mercure Szczyrk
    # Resort).
    proposals.sort(key=lambda p: p["score"], reverse=True)
    used_cols = set()
    for p in proposals:
        if any(c in used_cols for c in p["cols"]):
            log(f"  [pomijam - kolizja] pokoj '{p['room_name']}' - kolumna(y) zajete juz przez mocniej "
                f"pasujacy pokoj")
            continue
        mapping[p["room_id"]] = p["cols"]
        used_cols.update(p["cols"])
        multi_note = f" (zapisano do {len(p['cols'])} kolumn: {', '.join(p['cols'])})" if len(p["cols"]) > 1 else ""
        log(f"  Dopasowanie pokoju: '{p['room_name']}' -> '{p['names']}' [pewnosc {p['score']:.2f}]{multi_note}")

    return mapping


def _find_label_rows(ws, max_scan_row: int = 20) -> dict:
    labels = {}
    for r in range(1, max_scan_row + 1):
        val = ws[f"A{r}"].value
        if isinstance(val, str) and val.strip() in MANDALA_HEADER_LABELS:
            labels[val.strip()] = r
    return labels


def _read_file_room_columns(ws, room_row: int, board_row: int = None) -> list:
    cols = []
    for c in range(2, ws.max_column + 1):
        cell = ws.cell(row=room_row, column=c)
        if isinstance(cell.value, str) and cell.value.strip():
            board_text = None
            if board_row is not None:
                bcell = ws.cell(row=board_row, column=c)
                if isinstance(bcell.value, str):
                    board_text = bcell.value.strip()
            cols.append((cell.column_letter, cell.value.strip(), board_text))
    return cols


def write_mandala_xlsx_merge(
    rows: list,
    calendar_days: dict,
    base_file_path: str,
    output_path: str,
    best_price_only: bool = False,
    log=print,
) -> dict:
    """Scala wyniki scrapowania z JUZ ISTNIEJACYM plikiem cennika (np.
    biezacymi cenami z naszego systemu), zamiast zaczynac od pustego
    szablonu.

    Cala "gorna czesc" arkusza (MLOS, nazwy pokoi, Room/Board/Package ID,
    Internal Name...) pozostaje DOKLADNIE taka jak w pliku wejsciowym -
    modyfikowana jest tylko siatka cen (data x pokoj).

    - best_price_only=False (nadpisanie): kazda data/pokoj, dla ktorej
      scraper znalazl cene, nadpisuje istniejaca wartosc. Daty, ktorych
      scraper NIE znalazl (np. konkurencja niedostepna danego dnia albo
      poza sprawdzanym zakresem), zostaja bez zmian - te wiersze po prostu
      nie pojawiaja sie w `rows`, wiec nie sa dotykane.
    - best_price_only=True (najlepsza cena): nadpisuje istniejaca wartosc
      TYLKO gdy nowa (konkurencyjna) cena jest od niej nizsza, albo gdy
      komorka byla pusta. W przeciwnym razie wlasna (juz lepsza lub rowna)
      cena zostaje bez zmian.

    Zwraca dict:
    - "changes": lista zmian (tylko dla best_price_only=True) - dict z
      kluczami arrival_date/room_name/price_before/price_after dla kazdej
      komorki, ktora miala juz jakas cene i zostala nadpisana bo
      konkurencja byla tansza. Puste komorki (brak wczesniejszej ceny) nie
      licza sie jako "zmiana" - po prostu uzupelniamy brakujaca cene.
    - "total_rooms" / "matched_rooms" / "unmatched_names": statystyki
      dopasowania pokoi - przydatne np. do ostrzezenia uzytkownika, gdy
      dopasowanie nie powiodlo sie w ogole (plik wynikowy zawiera wtedy
      stare, niedotkniete ceny z przeslanego pliku).
    """
    if not rows:
        raise ValueError("Brak danych do zapisania (pusta lista wierszy)")

    wb = load_workbook(base_file_path)
    if MANDALA_SHEET not in wb.sheetnames:
        raise ValueError(
            f"Przeslany plik nie zawiera arkusza '{MANDALA_SHEET}' - sprawdz, czy to wlasciwy plik cennika"
        )
    ws = wb[MANDALA_SHEET]
    label_rows = _find_label_rows(ws)

    # Niektore hotele w Mandali maja w danym momencie aktywna WYLACZNIE oferte
    # promocyjna - wtedy arkusz "Regular" jest pustym szkieletem (sam
    # osierocony wiersz "date", bez wiersza "Room" i bez zadnych danych), a
    # realne ceny trafiaja do arkusza "Promotional" zamiast standardowego.
    # Zanim uznamy plik za nierozpoznany, sprawdzamy czy taki alternatywny
    # arkusz istnieje i ma poprawna strukture (sprawdzone na realnych
    # przypadkach z bulk testu: Grand Hotel Tiffi, Hotel Skalite SPA &
    # Wellness - oba maja identyczny uklad naglowkow w "Promotional", tylko
    # "Regular" jest u nich pusty):
    if "Room" not in label_rows and "Promotional" in wb.sheetnames and MANDALA_SHEET != "Promotional":
        promo_ws = wb["Promotional"]
        promo_labels = _find_label_rows(promo_ws)
        if "Room" in promo_labels and "date" in promo_labels:
            log(f"  [uwaga] arkusz '{MANDALA_SHEET}' jest pusty dla tego hotelu (widocznie ma aktywna "
                f"tylko oferte promocyjna) - uzywam arkusza 'Promotional' zamiast tego")
            ws = promo_ws
            label_rows = promo_labels

    room_row = label_rows.get("Room")
    board_row = label_rows.get("Board")
    date_row = label_rows.get("date")
    if room_row is None or date_row is None:
        raise ValueError(
            "Nie rozpoznano struktury przeslanego pliku (brak wiersza 'Room' lub 'date' w kolumnie A) - "
            "sprawdz, czy to plik w formacie Mandala"
        )
    data_start_row = date_row + 1

    file_rooms = _read_file_room_columns(ws, room_row, board_row)
    if not file_rooms:
        raise ValueError("Nie znaleziono zadnych kolumn z nazwami pokoi w przeslanym pliku")

    scraped_rooms = _unique_rooms_in_order(rows)
    meal_code = rows[0].get("meal_code") or None
    log(f"Dopasowuje {len(scraped_rooms)} zescrapowanych typow pokoi do {len(file_rooms)} kolumn w przeslanym pliku...")
    room_col = match_rooms(scraped_rooms, file_rooms, meal_code=meal_code, log=log)

    unmatched = [name for rid, name in scraped_rooms if rid not in room_col]
    if unmatched:
        log(f"  [UWAGA] nie dopasowano {len(unmatched)} pokoi - ich ceny NIE zostana zapisane: {', '.join(unmatched)}")
    if scraped_rooms and not room_col:
        log(
            "  [UWAGA KRYTYCZNA] Nie dopasowano ANI JEDNEGO pokoju do przeslanego pliku! "
            "Plik wynikowy zawiera STARE ceny z przeslanego pliku, NIE nowe ceny ze scrapowania. "
            "Sprawdz recznie nazwy pokoi w obu plikach."
        )

    end_date = max(datetime.strptime(r["arrival_date"], "%Y-%m-%d").date() for r in rows)
    date_map = _extend_date_rows(ws, [MANDALA_DATE_COL], data_start_row, end_date, WEEKEND_FILL_MANDALA, log=log)

    written = 0
    kept = 0
    skipped_out_of_range = 0
    changes = []
    for r in rows:
        cols = room_col.get(r["room_id"])
        if not cols:
            continue
        new_price = r["price_per_night"]
        if new_price is None:
            continue
        arrival = datetime.strptime(r["arrival_date"], "%Y-%m-%d").date()
        row_num = date_map.get(arrival)
        if row_num is None:
            skipped_out_of_range += 1
            continue

        for col in cols:
            cell = ws[f"{col}{row_num}"]
            existing = cell.value

            if best_price_only and isinstance(existing, (int, float)):
                if existing <= new_price:
                    kept += 1
                    continue
                changes.append({
                    "arrival_date": r["arrival_date"],
                    "room_name": r["room_name"],
                    "price_before": existing,
                    "price_after": new_price,
                })

            cell.value = new_price
            cell.number_format = "0.00"
            written += 1

    if skipped_out_of_range:
        log(f"  [UWAGA] {skipped_out_of_range} wpisow pominieto - ich data wypada przed poczatkiem przeslanego pliku")

    wb.save(output_path)
    mode_txt = "najlepsza cena (podmieniono tylko tam, gdzie konkurencja tansza)" if best_price_only else "nadpisanie"
    log(
        f"Zapisano scalony cennik (Mandala, tryb: {mode_txt}): {output_path} "
        f"({written} cen zaktualizowanych, {kept} pozostawionych bo wlasna cena juz najlepsza)"
    )
    return {
        "changes": changes,
        "total_rooms": len(scraped_rooms),
        "matched_rooms": len(room_col),
        "unmatched_names": unmatched,
    }


DIFF_REPORT_HEADERS = ["Data przyjazdu", "Pokoj", "Cena przed", "Cena po", "Roznica (zl)", "Roznica (%)"]


def write_price_diff_report(changes: list, hotel_name: str, output_path: str, log=print) -> None:
    """Zapisuje prosty raport zmian cen z trybu 'najlepsza cena' - dla
    kazdej daty/pokoju, gdzie konkurencja byla tansza i nasza cena zostala
    podmieniona, pokazuje cene przed i po zmianie oraz wyliczona roznice
    (wyroznona kolorem). To osobny, plaski plik xlsx - nie oparty na
    szablonie Mandala/Ava.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Zmiany cen"

    ws["A1"] = f"Raport zmian cen (najlepsza cena) - {hotel_name}"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = f"Liczba zmienionych cen: {len(changes)}"

    header_row = 4
    for col_idx, header in enumerate(DIFF_REPORT_HEADERS, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = Font(bold=True)

    highlight_fill = PatternFill(start_color="FFFCE4E4", end_color="FFFCE4E4", fill_type="solid")
    diff_font = Font(color="FFCC0000", bold=True)

    sorted_changes = sorted(changes, key=lambda c: (c["arrival_date"], c["room_name"]))
    for row_idx, change in enumerate(sorted_changes, start=header_row + 1):
        arrival = datetime.strptime(change["arrival_date"], "%Y-%m-%d")
        before = change["price_before"]
        after = change["price_after"]
        diff_val = round(before - after, 2)
        diff_pct = (diff_val / before) if before else None

        ws.cell(row=row_idx, column=1, value=arrival).number_format = "yyyy-mm-dd"
        ws.cell(row=row_idx, column=2, value=change["room_name"])
        ws.cell(row=row_idx, column=3, value=before).number_format = "0.00"
        ws.cell(row=row_idx, column=4, value=after).number_format = "0.00"

        diff_cell = ws.cell(row=row_idx, column=5, value=diff_val)
        diff_cell.number_format = "0.00"
        diff_cell.font = diff_font

        pct_cell = ws.cell(row=row_idx, column=6, value=diff_pct)
        pct_cell.number_format = "0.0%"
        pct_cell.font = diff_font

        for col_idx in range(1, len(DIFF_REPORT_HEADERS) + 1):
            ws.cell(row=row_idx, column=col_idx).fill = highlight_fill

    widths = [14, 40, 12, 12, 14, 12]
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=header_row, column=col_idx).column_letter].width = width

    wb.save(output_path)
    log(f"Zapisano raport zmian cen: {output_path} ({len(changes)} zmienionych cen)")


# ----------------------------------------------------------------------- Ava

def write_ava_xlsx(rows: list[dict], output_path: str, log=print, write_block: str = "current") -> None:
    """Zapisuje raport dostepnosci w formacie szablonu Ava.

    write_block: "current" - wpisuje dostepnosc w blok "Current availability"
        (kolumny C/D, pod naglowkiem B1). To domyslne, historyczne zachowanie
        (triverna_gui.py).
        "changes" - wpisuje dostepnosc w blok "Please make changes" (kolumny
        G/H, pod naglowkiem F1) - tak, jak tego oczekuje raport Eurotours
        AgentPlus (eurotours_report.py). Blok "Current availability" zostaje
        wtedy pusty (poza data w kolumnie B).

    UWAGA (dotyczy triverna.pl, write_block="current"): Triverna udostepnia
    publicznie tylko zbiorcza (dla calej oferty, nie per typ pokoju) liczbe
    dostepnych pokoi na dana date - ta sama wartosc trafia wiec do kolumny
    kazdego pokoju. Informacja o tym ograniczeniu jest dopisywana jako
    komentarz w pliku wynikowym.
    """
    if not rows:
        raise ValueError("Brak danych do zapisania (pusta lista wierszy)")
    if write_block not in ("current", "changes"):
        raise ValueError(f"Nieznana wartosc write_block: {write_block!r}")

    room_cols = AVA_ROOM_COLS if write_block == "current" else AVA_ROOM_COLS_CHANGES
    header_row = 2

    wb = load_workbook(AVA_TEMPLATE_PATH)
    ws = wb[AVA_SHEET]

    room_order = _unique_rooms_in_order(rows)
    if len(room_order) > len(room_cols):
        log(f"  [UWAGA] oferta ma {len(room_order)} typow pokoi - szablon Ava (dostepnosc) obsluguje max. "
            f"{len(room_cols)}. Eksportuje tylko pierwsze {len(room_cols)}.")
        room_order = room_order[: len(room_cols)]

    room_col: dict = {}
    for idx, (room_id, room_name) in enumerate(room_order):
        col = room_cols[idx]
        room_col[room_id] = col
        ws[f"{col}{header_row}"] = room_name

    for col in room_cols:
        if col not in room_col.values():
            ws[f"{col}{header_row}"] = None

    if write_block == "current":
        note = (
            "Triverna.pl udostepnia publicznie tylko zbiorcza (nie per-pokoj) liczbe "
            "dostepnych pokoi dla danej daty przyjazdu - ta sama wartosc jest wpisana "
            "we wszystkie kolumny pokoi ponizej."
        )
        ws["B1"].comment = Comment(note, "Triverna Scraper")

    _rebase_template_dates_from_today(ws, [AVA_DATE_COL, AVA_MIRROR_DATE_COL], AVA_DATA_START_ROW, WEEKEND_FILL_AVA)
    end_date = max(datetime.strptime(r["arrival_date"], "%Y-%m-%d").date() for r in rows)
    date_map = _extend_date_rows(
        ws, [AVA_DATE_COL, AVA_MIRROR_DATE_COL], AVA_DATA_START_ROW, end_date, WEEKEND_FILL_AVA, log=log
    )

    written = 0
    for r in rows:
        arrival = datetime.strptime(r["arrival_date"], "%Y-%m-%d").date()
        row_num = date_map.get(arrival)
        col = room_col.get(r["room_id"])
        if row_num is None or col is None:
            continue
        cell = ws[f"{col}{row_num}"]
        cell.value = r["available_quantity"]
        cell.number_format = "0"
        written += 1

    wb.save(output_path)
    log(f"Zapisano raport dostepnosci (Ava): {output_path} ({written} wpisow)")
