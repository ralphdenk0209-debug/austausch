// Playwright-Werkzeug – Hilfsbefehle für Node (gleiche Befehle wie helfer.py)
// Aufruf: node helfer.js <befehl> ...   (normalerweise über START.bat)
//   bilder   <adresse|liste.txt> <ordner>
//   pdf      <adresse|liste.txt> <ordner>
//   pruefen  <pruefen.txt> <ordner>
//   tabelle  <adresse> <ordner> [css-auswahl]
//   anmelden <adresse> <speicher.json>
// Umgebung: PW_BROWSER (msedge | chrome | chromium), PW_SPEICHER (gespeicherte Anmeldung)

let pw;
try { pw = require('playwright'); } catch (e) { pw = require('@playwright/test'); }
const fs = require('fs');
const path = require('path');
const readline = require('readline');

const BROWSER = process.env.PW_BROWSER || 'msedge';
const SPEICHER = process.env.PW_SPEICHER || '';

async function starten(sichtbar = false) {
  const opts = { headless: !sichtbar };
  if (BROWSER && BROWSER !== 'chromium') opts.channel = BROWSER;
  const browser = await pw.chromium.launch(opts);
  const ctxOpts = { viewport: { width: 1440, height: 900 }, locale: 'de-DE' };
  if (SPEICHER && fs.existsSync(SPEICHER)) ctxOpts.storageState = SPEICHER;
  const context = await browser.newContext(ctxOpts);
  return { browser, context };
}

function adresse(u) {
  u = u.trim();
  return /^[a-z]+:\/\//i.test(u) ? u : 'https://' + u;
}

function liesListe(eingabe) {
  if (fs.existsSync(eingabe)) {
    return fs.readFileSync(eingabe, 'utf8').replace(/^﻿/, '').split(/\r?\n/)
      .map(z => z.trim()).filter(z => z && !z.startsWith('#'));
  }
  return [eingabe];
}

function dateiname(url, nr) {
  let t;
  try { const u = new URL(url); t = u.hostname + u.pathname; } catch (e) { t = url; }
  t = t.replace(/[^A-Za-z0-9äöüÄÖÜß._-]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 70) || 'seite';
  return String(nr).padStart(2, '0') + '_' + t;
}

function csvZeile(werte) {
  return werte.map(w => {
    const s = (w === null || w === undefined) ? '' : String(w);
    return /[;"\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }).join(';');
}
function schreibeCsv(datei, zeilen) {
  // BOM + Semikolon, damit Excel (deutsch) Umlaute und Spalten richtig erkennt
  fs.writeFileSync(datei, '﻿' + zeilen.map(csvZeile).join('\r\n') + '\r\n', 'utf8');
}

async function laden(page, url) {
  const antwort = await page.goto(url, { waitUntil: 'load', timeout: 45000 });
  try { await page.waitForLoadState('networkidle', { timeout: 8000 }); } catch (e) { /* manche Seiten werden nie ruhig */ }
  return antwort;
}

async function bilderOderPdf(art, eingabe, ordner) {
  const { browser, context } = await starten();
  const page = await context.newPage();
  const liste = liesListe(eingabe);
  let nr = 0;
  for (const zeile of liste) {
    nr++;
    const url = adresse(zeile.split('|')[0]);
    try {
      await laden(page, url);
      const name = dateiname(url, nr);
      if (art === 'pdf') {
        await page.pdf({ path: path.join(ordner, name + '.pdf'), format: 'A4', printBackground: true });
      } else {
        await page.screenshot({ path: path.join(ordner, name + '.png'), fullPage: true });
      }
      console.log('OK      ' + url);
    } catch (e) {
      console.log('FEHLER  ' + url + '  ' + e.message.split('\n')[0]);
    }
  }
  await browser.close();
}

async function pruefen(datei, ordner) {
  if (!fs.existsSync(datei)) { console.log('Datei fehlt: ' + datei); return; }
  const { browser, context } = await starten();
  const page = await context.newPage();
  const zeilen = [['Adresse', 'HTTP-Status', 'Titel', 'Ladezeit (s)', 'Erwarteter Text', 'Text gefunden', 'Ergebnis', 'Fehler']];
  let nr = 0, fehler = 0;
  for (const zeile of liesListe(datei)) {
    nr++;
    const teile = zeile.split('|');
    const url = adresse(teile[0]);
    const text = (teile[1] || '').trim();
    const start = Date.now();
    let status = '', titel = '', gefunden = '', ok = false, meldung = '';
    try {
      const antwort = await laden(page, url);
      status = antwort ? antwort.status() : '';
      titel = await page.title();
      if (text) {
        const inhalt = await page.locator('body').innerText();
        gefunden = inhalt.toLowerCase().includes(text.toLowerCase()) ? 'ja' : 'nein';
      }
      ok = (!status || status < 400) && gefunden !== 'nein';
    } catch (e) {
      meldung = e.message.split('\n')[0];
    }
    const sek = ((Date.now() - start) / 1000).toFixed(1).replace('.', ',');
    if (!ok) {
      fehler++;
      try { await page.screenshot({ path: path.join(ordner, 'FEHLER_' + dateiname(url, nr) + '.png') }); } catch (e) { }
    }
    console.log((ok ? 'OK      ' : 'FEHLER  ') + url + (text ? '  [' + text + ': ' + (gefunden || '?') + ']' : '') + (meldung ? '  ' + meldung : ''));
    zeilen.push([url, status, titel, sek, text, gefunden, ok ? 'OK' : 'FEHLER', meldung]);
  }
  await browser.close();
  schreibeCsv(path.join(ordner, 'pruefung.csv'), zeilen);
  console.log('');
  console.log((nr - fehler) + ' von ' + nr + ' Seiten in Ordnung. Bericht: pruefung.csv');
}

async function tabelle(url, ordner, auswahl) {
  const { browser, context } = await starten();
  const page = await context.newPage();
  await laden(page, adresse(url));
  const tabellen = await page.$$eval(auswahl || 'table', els => els.map(el => {
    const zeilen = Array.from(el.querySelectorAll('tr'))
      .map(tr => Array.from(tr.querySelectorAll('th,td')).map(c => c.innerText.replace(/\s+/g, ' ').trim()))
      .filter(z => z.length);
    if (zeilen.length) return zeilen;
    // Keine echte Tabelle: jede Textzeile wird eine Zeile
    return el.innerText.split('\n').map(s => s.trim()).filter(Boolean).map(s => [s]);
  }));
  await browser.close();
  if (!tabellen.length) { console.log('Keine passende Tabelle gefunden.'); return; }
  tabellen.forEach((t, i) => {
    const datei = path.join(ordner, 'tabelle_' + (i + 1) + '.csv');
    schreibeCsv(datei, t);
    console.log('tabelle_' + (i + 1) + '.csv  (' + t.length + ' Zeilen)');
  });
}

async function anmelden(url, speicher) {
  const { browser, context } = await starten(true);
  const page = await context.newPage();
  await page.goto(adresse(url));
  console.log('Melde dich im Browser an. Danach hier Enter drücken.');
  await new Promise(fertig => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question('', () => { rl.close(); fertig(); });
  });
  await context.storageState({ path: speicher });
  await browser.close();
  console.log('Anmeldung gespeichert: ' + speicher);
}

(async () => {
  const [befehl, a, b, c] = process.argv.slice(2);
  try {
    if (befehl === 'bilder') await bilderOderPdf('bild', a, b);
    else if (befehl === 'pdf') await bilderOderPdf('pdf', a, b);
    else if (befehl === 'pruefen') await pruefen(a, b);
    else if (befehl === 'tabelle') await tabelle(a, b, c);
    else if (befehl === 'anmelden') await anmelden(a, b);
    else { console.log('Unbekannter Befehl: ' + befehl); process.exitCode = 2; }
  } catch (e) {
    console.log('Fehler: ' + e.message);
    if (/Executable doesn't exist|browserType.launch/i.test(e.message)) {
      console.log('Tipp: In einstellungen.json "Browser" auf "msedge" oder "chrome" stellen.');
    }
    process.exitCode = 1;
  }
})();
