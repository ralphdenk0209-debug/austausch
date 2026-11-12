# Playwright-Werkzeug – Hilfsbefehle für Python (gleiche Befehle wie helfer.js)
# Aufruf: python helfer.py <befehl> ...   (normalerweise über START.bat)
#   bilder   <adresse|liste.txt> <ordner>
#   pdf      <adresse|liste.txt> <ordner>
#   pruefen  <pruefen.txt> <ordner>
#   tabelle  <adresse> <ordner> [css-auswahl]
#   anmelden <adresse> <speicher.json>
# Umgebung: PW_BROWSER (msedge | chrome | chromium), PW_SPEICHER (gespeicherte Anmeldung)
import os
import re
import sys
import time
from urllib.parse import urlparse

from playwright.sync_api import sync_playwright

try:
    sys.stdout.reconfigure(errors='replace')
except Exception:
    pass

BROWSER = os.environ.get('PW_BROWSER') or 'msedge'
SPEICHER = os.environ.get('PW_SPEICHER') or ''


def starten(p, sichtbar=False):
    opts = {'headless': not sichtbar}
    if BROWSER and BROWSER != 'chromium':
        opts['channel'] = BROWSER
    browser = p.chromium.launch(**opts)
    ctx_opts = {'viewport': {'width': 1440, 'height': 900}, 'locale': 'de-DE'}
    if SPEICHER and os.path.exists(SPEICHER):
        ctx_opts['storage_state'] = SPEICHER
    return browser, browser.new_context(**ctx_opts)


def adresse(u):
    u = u.strip()
    return u if re.match(r'^[a-z]+://', u, re.I) else 'https://' + u


def lies_liste(eingabe):
    if os.path.exists(eingabe):
        with open(eingabe, encoding='utf-8-sig') as f:
            return [z.strip() for z in f if z.strip() and not z.strip().startswith('#')]
    return [eingabe]


def dateiname(url, nr):
    try:
        u = urlparse(url)
        t = (u.hostname or '') + (u.path or '')
    except Exception:
        t = url
    t = re.sub(r'[^A-Za-z0-9äöüÄÖÜß._-]+', '_', t).strip('_')[:70] or 'seite'
    return f'{nr:02d}_{t}'


def csv_zeile(werte):
    aus = []
    for w in werte:
        s = '' if w is None else str(w)
        if re.search(r'[;"\r\n]', s):
            s = '"' + s.replace('"', '""') + '"'
        aus.append(s)
    return ';'.join(aus)


def schreibe_csv(datei, zeilen):
    # BOM + Semikolon, damit Excel (deutsch) Umlaute und Spalten richtig erkennt
    with open(datei, 'w', encoding='utf-8-sig', newline='') as f:
        f.write('\r\n'.join(csv_zeile(z) for z in zeilen) + '\r\n')


def laden(page, url):
    antwort = page.goto(url, wait_until='load', timeout=45000)
    try:
        page.wait_for_load_state('networkidle', timeout=8000)
    except Exception:
        pass  # manche Seiten werden nie ruhig
    return antwort


def bilder_oder_pdf(art, eingabe, ordner):
    with sync_playwright() as p:
        browser, ctx = starten(p)
        page = ctx.new_page()
        for nr, zeile in enumerate(lies_liste(eingabe), 1):
            url = adresse(zeile.split('|')[0])
            try:
                laden(page, url)
                name = dateiname(url, nr)
                if art == 'pdf':
                    page.pdf(path=os.path.join(ordner, name + '.pdf'), format='A4', print_background=True)
                else:
                    page.screenshot(path=os.path.join(ordner, name + '.png'), full_page=True)
                print('OK      ' + url)
            except Exception as e:
                print('FEHLER  ' + url + '  ' + str(e).splitlines()[0])
        browser.close()


def pruefen(datei, ordner):
    if not os.path.exists(datei):
        print('Datei fehlt: ' + datei)
        return
    zeilen = [['Adresse', 'HTTP-Status', 'Titel', 'Ladezeit (s)', 'Erwarteter Text', 'Text gefunden', 'Ergebnis', 'Fehler']]
    fehler = 0
    liste = lies_liste(datei)
    with sync_playwright() as p:
        browser, ctx = starten(p)
        page = ctx.new_page()
        for nr, zeile in enumerate(liste, 1):
            teile = zeile.split('|')
            url = adresse(teile[0])
            text = teile[1].strip() if len(teile) > 1 else ''
            start = time.time()
            status, titel, gefunden, ok, meldung = '', '', '', False, ''
            try:
                antwort = laden(page, url)
                status = antwort.status if antwort else ''
                titel = page.title()
                if text:
                    inhalt = page.locator('body').inner_text()
                    gefunden = 'ja' if text.lower() in inhalt.lower() else 'nein'
                ok = (not status or status < 400) and gefunden != 'nein'
            except Exception as e:
                meldung = str(e).splitlines()[0]
            sek = f'{time.time() - start:.1f}'.replace('.', ',')
            if not ok:
                fehler += 1
                try:
                    page.screenshot(path=os.path.join(ordner, 'FEHLER_' + dateiname(url, nr) + '.png'))
                except Exception:
                    pass
            print(('OK      ' if ok else 'FEHLER  ') + url
                  + (f'  [{text}: {gefunden or "?"}]' if text else '') + (f'  {meldung}' if meldung else ''))
            zeilen.append([url, status, titel, sek, text, gefunden, 'OK' if ok else 'FEHLER', meldung])
        browser.close()
    schreibe_csv(os.path.join(ordner, 'pruefung.csv'), zeilen)
    print('')
    print(f'{len(liste) - fehler} von {len(liste)} Seiten in Ordnung. Bericht: pruefung.csv')


JS_TABELLEN = '''els => els.map(el => {
  const zeilen = Array.from(el.querySelectorAll('tr'))
    .map(tr => Array.from(tr.querySelectorAll('th,td')).map(c => c.innerText.replace(/\\s+/g, ' ').trim()))
    .filter(z => z.length);
  if (zeilen.length) return zeilen;
  return el.innerText.split('\\n').map(s => s.trim()).filter(Boolean).map(s => [s]);
})'''


def tabelle(url, ordner, auswahl=None):
    with sync_playwright() as p:
        browser, ctx = starten(p)
        page = ctx.new_page()
        laden(page, adresse(url))
        tabellen = page.eval_on_selector_all(auswahl or 'table', JS_TABELLEN)
        browser.close()
    if not tabellen:
        print('Keine passende Tabelle gefunden.')
        return
    for i, t in enumerate(tabellen, 1):
        schreibe_csv(os.path.join(ordner, f'tabelle_{i}.csv'), t)
        print(f'tabelle_{i}.csv  ({len(t)} Zeilen)')


def anmelden(url, speicher):
    with sync_playwright() as p:
        browser, ctx = starten(p, sichtbar=True)
        page = ctx.new_page()
        page.goto(adresse(url))
        input('Melde dich im Browser an. Danach hier Enter drücken.')
        ctx.storage_state(path=speicher)
        browser.close()
    print('Anmeldung gespeichert: ' + speicher)


def main():
    a = sys.argv[1:] + [None] * 4
    befehl = a[0]
    try:
        if befehl == 'bilder':
            bilder_oder_pdf('bild', a[1], a[2])
        elif befehl == 'pdf':
            bilder_oder_pdf('pdf', a[1], a[2])
        elif befehl == 'pruefen':
            pruefen(a[1], a[2])
        elif befehl == 'tabelle':
            tabelle(a[1], a[2], a[3])
        elif befehl == 'anmelden':
            anmelden(a[1], a[2])
        else:
            print(f'Unbekannter Befehl: {befehl}')
            sys.exit(2)
    except Exception as e:
        print('Fehler: ' + str(e))
        if re.search(r"Executable doesn't exist|BrowserType.launch", str(e), re.I):
            print('Tipp: In einstellungen.json "Browser" auf "msedge" oder "chrome" stellen.')
        sys.exit(1)


if __name__ == '__main__':
    main()
