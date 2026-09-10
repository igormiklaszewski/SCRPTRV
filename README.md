# SCRPTRV

Scraper cennika triverna.pl — wersja Streamlit (port z aplikacji desktopowej .exe).

## Uruchomienie lokalne

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Auto-pobieranie z Mandali

Funkcja automatycznego pobierania obecnego cennika z wewnętrznego systemu Mandala
(`mandala_fetcher.py`, Playwright + systemowy Chrome) wymaga dostępu do sieci
wewnętrznej firmy oraz zainstalowanego Chrome na maszynie, na której działa
aplikacja. Adres Mandali **nie jest zaszyty w kodzie** (repo jest publiczne) —
skonfiguruj go jako sekret:

- Lokalnie: utwórz `.streamlit/secrets.toml` (plik jest w `.gitignore`, nigdy
  nie trafia do repo) z zawartością:
  ```toml
  mandala_url = "http://<adres-wewnetrzny>/mandala"
  ```
- Na Streamlit Cloud: panel aplikacji → *Settings* → *Secrets*, ten sam klucz.

Gdy sekret nie jest ustawiony albo sieć wewnętrzna jest niedostępna (np.
Streamlit Community Cloud na czas testów), aplikacja automatycznie ukrywa opcję
auto-pobierania i pozwala wgrać plik z obecnym cennikiem ręcznie — po
przeniesieniu na docelowy serwer (współdzielony z Mandalą) i skonfigurowaniu
sekretu funkcja zacznie działać bez żadnych zmian w kodzie.

## Struktura

- `app.py` — interfejs Streamlit
- `triverna_scraper.py` — scraper cen (publiczne GraphQL API triverna.pl)
- `triverna_xlsx.py` — generowanie plików xlsx (Mandala/Ava)
- `mandala_fetcher.py` — automatyzacja pobierania z wewnętrznego systemu Mandala
- `templates/` — szablony xlsx
