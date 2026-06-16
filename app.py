from __future__ import annotations

import csv
import html
import io
import json
import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

import cgi
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


HOST = "127.0.0.1"
PORT = 8765
SETTINGS_FILE = Path(__file__).with_name("settings.json")

FIELDS = {
    "identifier": {
        "label": "Azonosító",
        "required": False,
        "help": "Ha üres, automatikusan: dátum + sorszám.",
    },
    "sender_account": {"label": "Feladó számlaszám", "required": True},
    "sender_name": {"label": "Feladó neve", "required": True},
    "beneficiary_account": {"label": "Címzett számlaszám", "required": True},
    "beneficiary_name": {"label": "Címzett neve", "required": True},
    "beneficiary_country": {"label": "Címzett országa", "required": False},
    "note": {"label": "Közlemény", "required": False},
    "document_no": {"label": "Bizonylatszám", "required": False},
    "legal_code": {"label": "Jogcímkód", "required": False},
    "currency": {"label": "Deviza", "required": True},
    "decimals": {"label": "Tizedesjegyek száma", "required": True},
    "amount": {"label": "Átutalandó összeg", "required": True},
    "value_date": {"label": "Valutanap", "required": True},
    "status": {"label": "Státusz", "required": False},
}

FX_FIELDS = {
    "identifier": {
        "label": "Azonosító",
        "required": False,
        "help": "Ha üres, automatikusan: dátum + sorszám.",
    },
    "sender_account": {"label": "Feladó számlaszám", "required": True},
    "sender_account_type": {"label": "Feladó számlaszám típusa", "required": True},
    "sender_name": {"label": "Feladó neve", "required": True},
    "beneficiary_account": {"label": "Címzett számlaszám", "required": True},
    "beneficiary_account_type": {"label": "Címzett számlaszám típusa", "required": True},
    "beneficiary_bank_name": {"label": "Címzett bank neve", "required": False},
    "beneficiary_name": {"label": "Címzett neve", "required": True},
    "beneficiary_country": {"label": "Bejegyzés országa", "required": True},
    "beneficiary_bank_swift": {"label": "Partner bank SWIFT kódja", "required": True},
    "correspondent_bank_1": {"label": "Levelező bank", "required": False},
    "swift_copy": {"label": "SWIFT másolat", "required": False},
    "fax_number": {"label": "Fax szám", "required": False},
    "custom_rate_use": {"label": "Egyedi árfolyam használata", "required": False},
    "custom_rate": {"label": "Egyedi árfolyam", "required": False},
    "bank_message": {"label": "Közlemény a bank számára", "required": False},
    "note": {"label": "Közlemény", "required": False},
    "legal_code": {"label": "Jogcímkód", "required": True},
    "permit_no": {"label": "Deviza engedély száma", "required": False},
    "permit_date": {"label": "Deviza engedély dátuma", "required": False},
    "urgent_use": {"label": "Sürgős teljesítés használata", "required": False},
    "urgent_execution": {"label": "Sürgős teljesítés", "required": False},
    "fax_copy_execution": {"label": "Fax másolat teljesítés", "required": False},
    "payout_currency": {"label": "Kifizetendő deviza", "required": True},
    "currency": {"label": "Átutalandó deviza", "required": True},
    "decimals": {"label": "Tizedesjegyek száma", "required": True},
    "amount": {"label": "Átutalandó összeg", "required": True},
    "value_date": {"label": "Feldolgozás kezdő dátuma", "required": False},
    "cost_bearer": {"label": "Költségviselés", "required": False},
    "partner_bank_country": {"label": "Partner bank országkódja", "required": True},
    "correspondent_bank_2": {"label": "Levelező bank 2. sor", "required": False},
    "correspondent_bank_3": {"label": "Levelező bank 3. sor", "required": False},
    "correspondent_bank_4": {"label": "Levelező bank 4. sor", "required": False},
}

TEMPLATE_EXAMPLE = {
    "identifier": "",
    "sender_account": "11111111-22222222-33333333",
    "sender_name": "Minta Feladó Kft",
    "beneficiary_account": "44444444-55555555-66666666",
    "beneficiary_name": "Minta Partner Kft",
    "beneficiary_country": "HU",
    "note": "Számla kiegyenlítés",
    "document_no": "BIZ001",
    "legal_code": "",
    "currency": "HUF",
    "decimals": "0",
    "amount": "12345",
    "value_date": "2026-06-17",
    "status": "",
}

FX_TEMPLATE_EXAMPLE = {
    "identifier": "",
    "sender_account": "HU00111111112222222233333333",
    "sender_account_type": "0",
    "sender_name": "Minta Feladó Kft",
    "beneficiary_account": "DE00123456781234567890",
    "beneficiary_account_type": "0",
    "beneficiary_bank_name": "Minta Bank",
    "beneficiary_name": "Minta Partner GmbH",
    "beneficiary_country": "DE",
    "beneficiary_bank_swift": "DEUTDEFF",
    "correspondent_bank_1": "",
    "swift_copy": "N",
    "fax_number": "",
    "custom_rate_use": "N",
    "custom_rate": "",
    "bank_message": "",
    "note": "Invoice payment",
    "legal_code": "1111",
    "permit_no": "",
    "permit_date": "",
    "urgent_use": "N",
    "urgent_execution": "",
    "fax_copy_execution": "",
    "payout_currency": "EUR",
    "currency": "EUR",
    "decimals": "2",
    "amount": "123.45",
    "value_date": "2026-06-17",
    "cost_bearer": "1",
    "partner_bank_country": "DE",
    "correspondent_bank_2": "",
    "correspondent_bank_3": "",
    "correspondent_bank_4": "",
}

BANKS = {
    "erste": {"label": "Erste Bank"},
}

FORMATS = {
    "erste_huf_payord": {
        "bank": "erste",
        "label": "Forint átutalás - PAYORD (DO)",
        "short_label": "PAYORD / DO",
        "badge": "PAYORD / DO · 941 karakter soronként CR/LF-fel",
        "description": "Erste forint átutalási import, 939 adatkarakter + CR/LF, tehát 941 karakter soronként.",
        "default_encoding": "cp1250",
        "line_length": 941,
        "fields": FIELDS,
        "template_example": TEMPLATE_EXAMPLE,
    },
    "erste_fx_payord": {
        "bank": "erste",
        "label": "Deviza átutalás - PAYORD (IN)",
        "short_label": "PAYORD / IN",
        "badge": "PAYORD / IN · 1048 karakter soronként CR/LF-fel",
        "description": "Erste deviza átutalási import, 1046 adatkarakter + CR/LF, tehát 1048 karakter soronként.",
        "default_encoding": "cp1250",
        "line_length": 1048,
        "fields": FX_FIELDS,
        "template_example": FX_TEMPLATE_EXAMPLE,
    },
}

ALIASES = {
    "identifier": ["azonosító", "azonosito", "id", "megbízás azonosító", "megbizas azonosito"],
    "sender_account": ["feladó számlaszám", "felado szamlaszam", "forrás számla", "forras szamla"],
    "sender_account_type": ["feladó számlaszám típusa", "felado szamlaszam tipusa", "feladó számlatípus"],
    "sender_name": ["feladó neve", "felado neve", "megbízó", "megbizo"],
    "beneficiary_account": [
        "címzett számlaszám",
        "cimzett szamlaszam",
        "kedvezményezett számlaszám",
        "kedvezmenyezett szamlaszam",
        "számlaszám",
        "szamlaszam",
    ],
    "beneficiary_account_type": ["címzett számlaszám típusa", "cimzett szamlaszam tipusa"],
    "beneficiary_bank_name": ["címzett bank neve", "cimzett bank neve", "partner bank neve", "bank neve"],
    "beneficiary_name": [
        "címzett neve",
        "cimzett neve",
        "kedvezményezett neve",
        "kedvezmenyezett neve",
        "partner neve",
    ],
    "beneficiary_country": ["ország", "orszag", "címzett országa", "cimzett orszaga"],
    "beneficiary_bank_swift": ["swift", "swift kód", "swift kod", "partner bank swift", "partner bank swift kódja"],
    "correspondent_bank_1": ["levelező bank", "levelezo bank"],
    "swift_copy": ["swift másolat", "swift masolat"],
    "fax_number": ["fax szám", "fax szam"],
    "custom_rate_use": ["egyedi árfolyam használata", "egyedi arfolyam hasznalata"],
    "custom_rate": ["egyedi árfolyam", "egyedi arfolyam"],
    "bank_message": ["közlemény a bank számára", "kozlemeny a bank szamara"],
    "note": ["közlemény", "kozlemeny", "megjegyzés", "megjegyzes"],
    "document_no": ["bizonylatszám", "bizonylatszam", "bizonylat", "sorszám", "sorszam"],
    "legal_code": ["jogcímkód", "jogcimkod", "jogcím", "jogcim"],
    "permit_no": ["deviza engedély száma", "deviza engedely szama"],
    "permit_date": ["deviza engedély dátuma", "deviza engedely datuma"],
    "urgent_use": ["sürgős teljesítés használata", "surgos teljesites hasznalata"],
    "urgent_execution": ["sürgős teljesítés", "surgos teljesites"],
    "fax_copy_execution": ["fax másolat teljesítés", "fax masolat teljesites"],
    "payout_currency": ["kifizetendő deviza", "kifizetendo deviza"],
    "currency": ["deviza", "valuta", "currency", "pénznem", "penznem"],
    "decimals": ["tizedesjegyek száma", "tizedesjegy", "decimals"],
    "amount": ["összeg", "osszeg", "átutalandó összeg", "atutalando osszeg", "amount"],
    "value_date": ["valutanap", "értéknap", "erteknap", "dátum", "datum"],
    "cost_bearer": ["költségviselés", "koltsegviseles"],
    "partner_bank_country": ["partner bank országkódja", "partner bank orszagkodja"],
    "correspondent_bank_2": ["levelező bank 2", "levelezo bank 2", "levelező bank 2. sor"],
    "correspondent_bank_3": ["levelező bank 3", "levelezo bank 3", "levelező bank 3. sor"],
    "correspondent_bank_4": ["levelező bank 4", "levelezo bank 4", "levelező bank 4. sor"],
    "status": ["státusz", "statusz", "status"],
}

HTML_PAGE = r"""<!doctype html>
<html lang="hu">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Banki TXT konverter</title>
  <style>
    :root {
      --bg: #f3f5f8;
      --surface: #ffffff;
      --surface-2: #f8fafc;
      --ink: #162033;
      --muted: #64748b;
      --line: #d9e1ec;
      --line-strong: #c9d3df;
      --accent: #bd2234;
      --accent-dark: #941827;
      --accent-soft: #fff1f3;
      --accent-2: #0f766e;
      --accent-2-soft: #edf9f6;
      --warn: #9a6400;
      --bad: #b42331;
      --bad-soft: #fff3f4;
      --shadow: 0 18px 44px rgba(17, 24, 39, 0.08);
      --shadow-soft: 0 8px 24px rgba(17, 24, 39, 0.06);
      font-family: Inter, Segoe UI, system-ui, -apple-system, sans-serif;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      color: var(--ink);
      background: var(--bg);
      font-size: 14px;
    }
    .shell { width: 100%; max-width: none; margin: 0; padding: 22px clamp(16px, 3vw, 40px) 36px; }
    header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 20px;
      padding: 4px 0 22px;
    }
    h1 { margin: 0 0 6px; font-size: 30px; letter-spacing: 0; line-height: 1.08; }
    p { margin: 0; color: var(--muted); line-height: 1.45; }
    .eyebrow {
      color: var(--accent);
      font-weight: 800;
      font-size: 12px;
      letter-spacing: .08em;
      text-transform: uppercase;
      margin-bottom: 7px;
    }
    .badge {
      border: 1px solid var(--line);
      border-radius: 6px;
      background: #fff;
      padding: 10px 12px;
      color: #475569;
      font-size: 13px;
      font-weight: 650;
      white-space: nowrap;
      box-shadow: var(--shadow-soft);
    }
    .layout { display: grid; gap: 18px; align-items: start; }
    .panel {
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      overflow: hidden;
    }
    .commandbar {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 12px;
      align-items: center;
      margin-bottom: 18px;
    }
    .command-summary {
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
      padding: 12px 14px;
      min-height: 52px;
      display: flex;
      align-items: center;
      gap: 12px;
      box-shadow: var(--shadow-soft);
    }
    .command-summary strong { font-size: 14px; }
    .command-summary span { color: var(--muted); font-size: 12px; }
    .command-summary span[data-kind="ok"] { color: var(--accent-2); }
    .command-summary span[data-kind="warn"] { color: var(--warn); }
    .command-summary span[data-kind="bad"] { color: var(--bad); }
    .command-actions {
      display: grid;
      grid-template-columns: repeat(4, auto);
      gap: 8px;
      align-items: center;
    }
    .import-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(220px, 1fr));
      gap: 12px;
      align-items: end;
      margin-bottom: 16px;
    }
    .import-grid .full-row { grid-column: 1 / -1; }
    .import-actions {
      display: flex;
      gap: 8px;
      align-items: center;
      flex-wrap: wrap;
      margin: 6px 0 18px;
    }
    .import-section-title {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: center;
      padding-top: 14px;
      border-top: 1px solid #edf0f3;
      margin-top: 6px;
    }
    .import-section-title h3 { margin: 0; font-size: 14px; }
    .panel h2 {
      margin: 0;
      padding: 14px 16px;
      font-size: 14px;
      letter-spacing: .01em;
      border-bottom: 1px solid var(--line);
      background: var(--surface-2);
    }
    .panel-body { padding: 16px; }
    label { display: block; font-weight: 750; font-size: 12px; margin-bottom: 6px; color: #263449; }
    input, select, button, textarea {
      font: inherit;
      color: var(--ink);
    }
    input[type="file"], input[type="text"], input[type="number"], input[type="date"], select {
      width: 100%;
      border: 1px solid var(--line);
      background: #fff;
      border-radius: 6px;
      padding: 9px 10px;
      min-height: 40px;
      font-size: 14px;
      outline: none;
      transition: border-color .15s ease, box-shadow .15s ease, background .15s ease;
    }
    input:focus, select:focus {
      border-color: rgba(15, 118, 110, .55);
      box-shadow: 0 0 0 3px rgba(15, 118, 110, .12);
    }
    input[type="file"] {
      padding: 6px;
      color: var(--muted);
      background: var(--surface-2);
    }
    input[type="file"]::file-selector-button {
      border: 0;
      border-radius: 5px;
      background: #263449;
      color: #fff;
      padding: 8px 10px;
      margin-right: 10px;
      font-weight: 750;
      cursor: pointer;
    }
    .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    .format-stack { display: grid; grid-template-columns: 1fr; gap: 10px; }
    .stack { display: grid; gap: 12px; }
    .hint { font-size: 11px; color: var(--muted); margin-top: 6px; line-height: 1.4; }
    .field-note { display: none; }
    .intro-note {
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 10px 12px;
      color: var(--muted);
      background: #fbfcfd;
      font-size: 13px;
      line-height: 1.45;
    }
    .actions { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-top: 12px; }
    button, .button-link {
      border: 1px solid transparent;
      border-radius: 6px;
      padding: 9px 11px;
      font-weight: 800;
      min-height: 40px;
      cursor: pointer;
      background: #eef2f5;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      text-decoration: none;
      transition: transform .12s ease, box-shadow .12s ease, border-color .12s ease, background .12s ease;
    }
    button:hover:not(:disabled), .button-link:hover { transform: translateY(-1px); box-shadow: var(--shadow-soft); }
    button.primary { background: var(--accent); color: #fff; border-color: var(--accent); }
    button.primary:hover:not(:disabled) { background: var(--accent-dark); }
    button.secondary, .button-link.secondary { background: #fff; border-color: var(--line-strong); color: var(--ink); }
    button.ghost, .button-link.ghost { background: var(--surface-2); border-color: var(--line); color: #334155; }
    button:disabled { cursor: not-allowed; opacity: .45; transform: none; box-shadow: none; }
    .mapping {
      display: grid;
      grid-template-columns: minmax(190px, 250px) minmax(180px, 1fr) minmax(160px, 230px);
      gap: 10px;
      align-items: center;
      padding: 10px 0;
      border-bottom: 1px solid #edf0f3;
    }
    .mapping:last-child { border-bottom: 0; }
    .req { color: var(--accent); font-weight: 800; }
    .map-help { color: var(--muted); font-size: 12px; }
    .default-input { min-width: 0; }
    .status {
      margin-top: 12px;
      border-radius: 6px;
      border: 1px solid var(--line);
      padding: 11px 12px;
      font-size: 12px;
      color: var(--muted);
      background: var(--surface-2);
      white-space: pre-wrap;
    }
    .status.ok { border-color: rgba(13,107,87,.35); color: var(--accent-2); background: #f2fbf7; }
    .status.warn { border-color: rgba(148,98,0,.35); color: var(--warn); background: #fff9ec; }
    .status.bad { border-color: rgba(165,29,42,.35); color: var(--bad); background: #fff4f5; }
    .sample-wrap { overflow: auto; border: 1px solid var(--line); border-radius: 6px; background:#fff; }
    table { border-collapse: collapse; width: 100%; min-width: 680px; font-size: 13px; }
    th, td { border-bottom: 1px solid #edf0f3; padding: 9px 10px; text-align: left; vertical-align: top; }
    th { background: #fbfcfd; font-weight: 750; color: #384554; position: sticky; top: 0; }
    td { color: #394757; }
    pre {
      margin: 0;
      max-height: 180px;
      overflow: auto;
      white-space: pre;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 12px;
      background: #101923;
      color: #f3f7fb;
      font-size: 12px;
      min-height: 58px;
    }
    .spec {
      display: none;
      grid-template-columns: repeat(3, 1fr);
      gap: 10px;
      margin-top: 13px;
      font-size: 12px;
      color: var(--muted);
    }
    .spec div { border: 1px solid var(--line); border-radius: 6px; padding: 9px; background: #fbfcfd; }
    .compact-actions { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
    .wide-action { grid-column: 1 / -1; }
    .result-grid { display: grid; grid-template-columns: repeat(3, minmax(160px, 1fr)); gap: 10px; }
    .metric {
      border: 1px solid var(--line);
      border-radius: 8px;
      background: linear-gradient(180deg, #fff 0%, var(--surface-2) 100%);
      padding: 13px;
      min-height: 74px;
      display: flex;
      flex-direction: column;
      justify-content: center;
      border-top: 3px solid var(--accent-2);
    }
    .metric strong { display: block; font-size: 22px; margin-bottom: 3px; line-height: 1.05; }
    .metric span { color: var(--muted); font-size: 12px; font-weight: 650; }
    .error-list { display: grid; gap: 8px; }
    .error-item { border: 1px solid rgba(165,29,42,.25); border-radius: 6px; padding: 9px; background: #fff4f5; color: var(--bad); font-size: 13px; }
    dialog {
      width: min(1120px, calc(100vw - 28px));
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      padding: 0;
    }
    dialog::backdrop { background: rgba(17,24,32,.34); }
    .dialog-head { display:flex; justify-content:space-between; align-items:center; gap:12px; padding:14px 16px; border-bottom:1px solid var(--line); background:var(--surface-2); }
    .dialog-head h2 { margin:0; font-size:16px; }
    .dialog-body { padding:16px; max-height:70vh; overflow:auto; }
    .help-grid { display:grid; gap:12px; }
    .help-grid section { border-bottom:1px solid #edf0f3; padding-bottom:12px; }
    .help-grid h3 { margin:0 0 6px; font-size:14px; }
    .help-grid p { font-size:13px; }
    .mapping-toolbar { display:flex; justify-content:space-between; gap:10px; align-items:center; margin-bottom:12px; }
    .muted-small { color: var(--muted); font-size: 12px; }
    @media (max-width: 880px) {
      .shell { padding: 16px 12px 28px; }
      .commandbar { grid-template-columns: 1fr; }
      .command-actions { grid-template-columns: 1fr 1fr; }
      .command-actions .wide-action { grid-column: 1 / -1; }
      .import-grid { grid-template-columns: 1fr; }
      header { display: block; }
      .badge { display: inline-block; margin-top: 12px; }
      .mapping { grid-template-columns: 1fr; gap: 7px; }
      .spec { grid-template-columns: 1fr; }
      .result-grid { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <div class="shell">
    <header>
      <div>
        <div class="eyebrow">Erste EDIFACT import</div>
        <h1>Banki TXT konverter</h1>
        <p>Excel feltöltése, banki formátum választása, majd fix hosszúságú importfájl letöltése.</p>
      </div>
      <div id="formatBadge" class="badge">PAYORD / DO · 941 karakter soronként CR/LF-fel</div>
    </header>

    <section class="commandbar">
      <div class="command-summary">
        <strong id="commandFormatName">PAYORD / DO</strong>
        <span id="statusBox">Nyisd meg az Import menüt, válassz fájlt, majd olvasd be.</span>
      </div>
      <div class="command-actions">
        <button id="openImportBtn" class="secondary" type="button">Import</button>
        <button id="convertBtn" class="primary" disabled>TXT letöltése</button>
        <a id="templateLink" class="button-link ghost" href="/template.xlsx" download>Excel sablon</a>
        <button id="helpBtn" class="ghost" type="button">? Súgó</button>
      </div>
    </section>

    <div class="layout">
      <main class="panel">
        <h2>Beolvasás eredménye</h2>
        <div class="panel-body">
          <div id="resultSummary" class="result-grid">
            <div class="metric"><strong>-</strong><span>Fejléc</span></div>
            <div class="metric"><strong>-</strong><span>Adatsor</span></div>
            <div class="metric"><strong>-</strong><span>Formátum</span></div>
          </div>
          <div id="errorArea" class="error-list" style="margin-top:14px;"></div>
          <div style="margin-top:14px;">
            <h3 style="font-size:14px;margin:0 0 8px;">Beolvasott minta</h3>
            <div id="sampleArea" class="sample-wrap"><table><tbody><tr><td>Nincs minta.</td></tr></tbody></table></div>
          </div>
        </div>
      </main>
    </div>

    <section class="panel" style="margin-top:16px;">
      <h2>TXT előnézet</h2>
      <div class="panel-body stack">
        <pre id="previewBox">A sikeres konvertálás után itt látszik az első rekord eleje.</pre>
      </div>
    </section>
  </div>

  <dialog id="importDialog">
    <div class="dialog-head">
      <h2>Import</h2>
      <button id="closeImportBtn" class="secondary" type="button">Bezárás</button>
    </div>
    <div class="dialog-body">
      <div class="import-grid">
        <div>
          <label for="bankSelect">Bank</label>
          <select id="bankSelect"></select>
        </div>
        <div>
          <label for="formatSelect">Formátum</label>
          <select id="formatSelect"></select>
          <div id="formatHelp" class="hint">A bank által elvárt TXT rekordforma.</div>
        </div>
        <div class="full-row">
          <label for="fileInput">Fájl</label>
          <input id="fileInput" type="file" accept=".xlsx,.xlsm,.csv">
        </div>
        <div>
          <label for="encoding">TXT kódolás</label>
          <select id="encoding">
            <option value="cp1250" selected>windows-1250</option>
            <option value="utf-8">UTF-8</option>
          </select>
        </div>
        <div>
          <label for="identifierDate">Azonosító dátuma</label>
          <input id="identifierDate" type="date">
        </div>
      </div>
      <div class="import-actions">
        <button id="inspectBtn" class="primary">Beolvasás</button>
        <a id="templateLinkInDialog" class="button-link ghost" href="/template.xlsx" download>Excel sablon</a>
      </div>
      <div class="import-section-title">
        <div>
          <h3>Oszlopok és alapértékek</h3>
          <div class="muted-small">Automatikusan mentődik a kiválasztott formátumhoz.</div>
        </div>
        <button id="useGuessesBtn" class="secondary" type="button">Automatikus kitöltés</button>
      </div>
      <div id="mappingArea" class="stack">
        <p>Még nincs beolvasott fejléc.</p>
      </div>
    </div>
  </dialog>

  <dialog id="helpDialog">
    <div class="dialog-head">
      <h2>Súgó</h2>
      <button id="closeHelpBtn" class="secondary" type="button">Bezárás</button>
    </div>
    <div class="dialog-body">
      <div class="help-grid">
        <section><h3>Bank és formátum</h3><p>A bank kiválasztása szűri a választható import formátumokat. Most az Erste EDIFACT forint PAYORD (DO) és deviza PAYORD (IN) formátumai érhetők el.</p></section>
        <section><h3>TXT kódolás</h3><p>A letöltött TXT karakterkészlete. Magyar banki importnál általában a windows-1250 a jó alapérték.</p></section>
        <section><h3>Azonosító dátuma</h3><p>Ha az Excelben nincs 14 számjegyű azonosító, a konverter ebből generál azonosítót: ÉÉÉÉHHNN + 6 jegyű sorszám.</p></section>
        <section><h3>Oszlopok és alapértékek</h3><p>A Beállítások alatt adható meg, hogy az Excel melyik oszlopa melyik banki mezőnek felel meg. Az alapérték akkor hasznos, ha egy adat minden sorban ugyanaz.</p></section>
        <section><h3>Hibakezelés</h3><p>Konvertálás előtt az app ellenőrzi a kötelező mezőket, dátumokat, összegeket és a mezőhosszokat. A hibák a Beolvasás eredménye panelen jelennek meg.</p></section>
      </div>
    </div>
  </dialog>

<script>
const BANKS = __BANKS_JSON__;
const FORMATS = __FORMATS_JSON__;
const FIELDS = __FIELDS_JSON__;
let currentInspect = null;
let currentSettings = { active_bank: "erste", active_format: "erste_huf_payord", formats: {} };
let saveTimer = null;

const el = id => document.getElementById(id);
const today = new Date();
el("identifierDate").value = today.toISOString().slice(0, 10);

function populateBanks() {
  const bankSelect = el("bankSelect");
  bankSelect.innerHTML = "";
  for (const [id, bank] of Object.entries(BANKS)) {
    const option = document.createElement("option");
    option.value = id;
    option.textContent = bank.label;
    bankSelect.appendChild(option);
  }
}

function populateFormats() {
  const bankId = el("bankSelect").value;
  const desired = currentSettings.active_format || el("formatSelect").value;
  const formatSelect = el("formatSelect");
  formatSelect.innerHTML = "";
  for (const [id, format] of Object.entries(FORMATS)) {
    if (format.bank !== bankId) continue;
    const option = document.createElement("option");
    option.value = id;
    option.textContent = format.label;
    formatSelect.appendChild(option);
  }
  if ([...formatSelect.options].some(o => o.value === desired)) {
    formatSelect.value = desired;
  }
  applySelectedFormat();
}

function selectedFormat() {
  return FORMATS[el("formatSelect").value] || Object.values(FORMATS)[0];
}

function currentFields() {
  return selectedFormat()?.fields || FIELDS;
}

function applySelectedFormat() {
  const format = selectedFormat();
  if (!format) return;
  el("formatBadge").textContent = format.badge || format.short_label || format.label;
  el("commandFormatName").textContent = format.short_label || format.label;
  el("formatHelp").textContent = format.description;
  el("templateLink").href = `/template.xlsx?format=${encodeURIComponent(el("formatSelect").value)}`;
  el("templateLinkInDialog").href = `/template.xlsx?format=${encodeURIComponent(el("formatSelect").value)}`;
  const settings = formatSettings();
  el("encoding").value = settings.encoding || format.default_encoding || "cp1250";
  if (settings.identifier_date) el("identifierDate").value = settings.identifier_date;
  buildMappings();
  saveSettingsDebounced();
}

function formatSettings(formatId = el("formatSelect").value) {
  currentSettings.formats ||= {};
  currentSettings.formats[formatId] ||= { mapping: {}, defaults: {} };
  currentSettings.formats[formatId].mapping ||= {};
  currentSettings.formats[formatId].defaults ||= {};
  return currentSettings.formats[formatId];
}

async function loadSettings() {
  try {
    const res = await fetch("/api/settings");
    if (res.ok) currentSettings = await res.json();
  } catch {}
}

function saveSettingsDebounced() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveSettings, 350);
}

async function saveSettings() {
  const formatId = el("formatSelect").value || "erste_huf_payord";
  currentSettings.active_bank = el("bankSelect").value || "erste";
  currentSettings.active_format = formatId;
  currentSettings.formats ||= {};
  const existing = formatSettings(formatId);
  const collected = collectConfig();
  if (!document.querySelector("[data-map]")) {
    collected.mapping = existing.mapping || {};
    collected.defaults = existing.defaults || {};
  }
  currentSettings.formats[formatId] = collected;
  try {
    await fetch("/api/settings", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(currentSettings)
    });
  } catch {
    setStatus("A beállítás mentése nem sikerült.", "warn");
  }
}

function setStatus(text, kind = "") {
  const box = el("statusBox");
  box.textContent = text;
  box.dataset.kind = kind;
}

function normalise(s) {
  return String(s || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function optionList(headers, selected) {
  let out = `<option value="">-- nincs oszlop --</option>`;
  for (const h of headers) {
    const safe = escapeHtml(h);
    out += `<option value="${safe}" ${h === selected ? "selected" : ""}>${safe}</option>`;
  }
  return out;
}

function escapeHtml(v) {
  return String(v ?? "").replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[c]));
}

function buildMappings() {
  const headers = currentInspect?.headers || [];
  const guesses = currentInspect?.guesses || {};
  const fields = currentFields();
  const settings = formatSettings();
  const area = el("mappingArea");
  area.innerHTML = "";
  for (const [key, info] of Object.entries(fields)) {
    const row = document.createElement("div");
    row.className = "mapping";
    const required = info.required ? `<span class="req">*</span>` : "";
    const guessed = settings.mapping[key] || guesses[key] || "";
    const defaultValue = settings.defaults[key] ?? defaultFor(key);
    row.innerHTML = `
      <div>
        <strong>${escapeHtml(info.label)} ${required}</strong>
        <div class="map-help">${escapeHtml(info.help || "Oszlopból vagy alapértékből jöhet.")}</div>
      </div>
      <select data-map="${escapeHtml(key)}">${optionList(headers, guessed)}</select>
      <input class="default-input" data-default="${escapeHtml(key)}" type="text"
        value="${escapeHtml(defaultValue)}" placeholder="Alapérték, ha nincs oszlop">
    `;
    area.appendChild(row);
  }
}

function defaultFor(key) {
  const defaults = {
    currency: "HUF",
    payout_currency: "EUR",
    decimals: "0",
    status: "",
    beneficiary_country: "HU",
    sender_account_type: "0",
    beneficiary_account_type: "0",
    swift_copy: "N",
    custom_rate_use: "N",
    urgent_use: "N",
    cost_bearer: "1",
    document_no: "",
    legal_code: ""
  };
  return defaults[key] ?? "";
}

function renderSample() {
  const headers = currentInspect?.headers || [];
  const rows = currentInspect?.sample || [];
  if (!headers.length) {
    el("sampleArea").innerHTML = `<table><tbody><tr><td>Nincs minta.</td></tr></tbody></table>`;
    return;
  }
  let html = "<table><thead><tr>";
  for (const h of headers) html += `<th>${escapeHtml(h)}</th>`;
  html += "</tr></thead><tbody>";
  for (const r of rows) {
    html += "<tr>";
    for (let i = 0; i < headers.length; i++) html += `<td>${escapeHtml(r[i] ?? "")}</td>`;
    html += "</tr>";
  }
  html += "</tbody></table>";
  el("sampleArea").innerHTML = html;
}

async function inspectFile() {
  const file = el("fileInput").files[0];
  if (!file) {
    setStatus("Előbb válassz ki egy fájlt.", "warn");
    return;
  }
  const form = new FormData();
  form.append("file", file);
  setStatus("Beolvasás...");
  const res = await fetch("/api/inspect", { method: "POST", body: form });
  const data = await res.json();
  if (!res.ok) {
    setStatus(data.error || "Nem sikerült beolvasni a fájlt.", "bad");
    renderErrors(data.errors || data.error || []);
    return;
  }
  currentInspect = data;
  buildMappings();
  renderSample();
  el("convertBtn").disabled = false;
  renderResultSummary(data);
  renderErrors([]);
  setStatus(`${data.headers.length} fejléc beolvasva, ${data.data_rows} adatsor észlelve.`, "ok");
  saveSettingsDebounced();
}

function renderResultSummary(data = null) {
  const format = selectedFormat();
  const rows = [
    [data?.headers?.length ?? "-", "Fejléc"],
    [data?.data_rows ?? "-", "Adatsor"],
    [format?.short_label ?? "-", "Formátum"]
  ];
  el("resultSummary").innerHTML = rows.map(([value, label]) =>
    `<div class="metric"><strong>${escapeHtml(value)}</strong><span>${escapeHtml(label)}</span></div>`
  ).join("");
}

function renderErrors(errors) {
  const list = Array.isArray(errors) ? errors : String(errors || "").split("\n").filter(Boolean);
  el("errorArea").innerHTML = list.map(e => `<div class="error-item">${escapeHtml(e)}</div>`).join("");
}

function collectConfig() {
  const mapping = {};
  const defaults = {};
  document.querySelectorAll("[data-map]").forEach(sel => mapping[sel.dataset.map] = sel.value);
  document.querySelectorAll("[data-default]").forEach(inp => defaults[inp.dataset.default] = inp.value);
  return {
    bank: el("bankSelect").value || "",
    format: el("formatSelect").value || "",
    encoding: el("encoding").value,
    identifier_date: el("identifierDate").value,
    mapping,
    defaults
  };
}

async function convertFile() {
  const file = el("fileInput").files[0];
  if (!file || !currentInspect) {
    setStatus("Előbb olvasd be a fájlt.", "warn");
    return;
  }
  const form = new FormData();
  form.append("file", file);
  form.append("config", JSON.stringify(collectConfig()));
  setStatus("Konvertálás...");
  const res = await fetch("/api/convert", { method: "POST", body: form });
  const contentType = res.headers.get("content-type") || "";
  if (!res.ok) {
    const data = contentType.includes("application/json") ? await res.json() : {error: await res.text()};
    setStatus(data.error || "Nem sikerült konvertálni.", "bad");
    renderErrors(data.errors || data.error || []);
    return;
  }
  const buffer = await res.arrayBuffer();
  const blob = new Blob([buffer], { type: contentType || "text/plain" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = res.headers.get("x-filename") || "payord_import.txt";
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);

  const decoderLabel = el("encoding").value === "cp1250" ? "windows-1250" : "utf-8";
  let text = "";
  try { text = new TextDecoder(decoderLabel).decode(buffer); }
  catch { text = new TextDecoder("utf-8").decode(buffer); }
  el("previewBox").textContent = text.slice(0, 700).replace(/\r/g, "\\r").replace(/\n/g, "\\n\n");
  const rows = res.headers.get("x-record-count") || "?";
  setStatus(`${rows} rekord elkészült és letöltődött.`, "ok");
  renderErrors([]);
  saveSettingsDebounced();
}

el("inspectBtn").addEventListener("click", () => inspectFile().catch(err => setStatus(err.message, "bad")));
el("convertBtn").addEventListener("click", () => convertFile().catch(err => setStatus(err.message, "bad")));
el("bankSelect").addEventListener("change", populateFormats);
el("formatSelect").addEventListener("change", applySelectedFormat);
el("openImportBtn").addEventListener("click", () => el("importDialog").showModal());
el("helpBtn").addEventListener("click", () => el("helpDialog").showModal());
el("closeImportBtn").addEventListener("click", () => el("importDialog").close());
el("closeHelpBtn").addEventListener("click", () => el("helpDialog").close());
el("useGuessesBtn").addEventListener("click", () => {
  const settings = formatSettings();
  settings.mapping = {...(currentInspect?.guesses || {})};
  buildMappings();
  saveSettingsDebounced();
});
el("mappingArea").addEventListener("change", saveSettingsDebounced);
el("mappingArea").addEventListener("input", saveSettingsDebounced);
el("encoding").addEventListener("change", saveSettingsDebounced);
el("identifierDate").addEventListener("change", saveSettingsDebounced);

(async function initApp() {
  populateBanks();
  await loadSettings();
  if (BANKS[currentSettings.active_bank]) el("bankSelect").value = currentSettings.active_bank;
  populateFormats();
  if (FORMATS[currentSettings.active_format]) el("formatSelect").value = currentSettings.active_format;
  applySelectedFormat();
  renderResultSummary();
})();
</script>
</body>
</html>
"""


def json_response(handler: BaseHTTPRequestHandler, data: dict[str, Any], status: int = 200) -> None:
    body = json.dumps(data, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def text_response(handler: BaseHTTPRequestHandler, body: str, status: int = 200) -> None:
    data = body.encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "text/html; charset=utf-8")
    handler.send_header("Content-Length", str(len(data)))
    handler.end_headers()
    handler.wfile.write(data)


def get_upload(handler: BaseHTTPRequestHandler) -> cgi.FieldStorage:
    return cgi.FieldStorage(
        fp=handler.rfile,
        headers=handler.headers,
        environ={
            "REQUEST_METHOD": "POST",
            "CONTENT_TYPE": handler.headers.get("Content-Type", ""),
            "CONTENT_LENGTH": handler.headers.get("Content-Length", "0"),
        },
    )


def field_value(form: cgi.FieldStorage, name: str, default: str = "") -> str:
    item = form[name] if name in form else None
    if item is None:
        return default
    if isinstance(item, list):
        item = item[0]
    if getattr(item, "filename", None):
        return default
    return str(item.value or default)


def upload_bytes(form: cgi.FieldStorage) -> tuple[str, bytes]:
    if "file" not in form:
        raise ValueError("Hiányzik a feltöltött fájl.")
    item = form["file"]
    filename = Path(item.filename or "upload.xlsx").name
    data = item.file.read()
    if not data:
        raise ValueError("A feltöltött fájl üres.")
    return filename, data


def clean_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y%m%d")
    if isinstance(value, date):
        return value.strftime("%Y%m%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def read_csv(data: bytes) -> tuple[list[str], list[list[Any]], str]:
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("cp1250")
    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=";,	,")
        rows = list(csv.reader(io.StringIO(text), dialect))
    except csv.Error:
        rows = list(csv.reader(io.StringIO(text), delimiter=";"))
    return ["CSV"], rows, "CSV"


def workbook_rows(data: bytes, filename: str) -> tuple[list[str], str, list[str], list[list[Any]]]:
    header_row = 1
    suffix = Path(filename).suffix.lower()
    if suffix == ".csv":
        sheets, rows, selected = read_csv(data)
        return sheets, selected, header_from_rows(rows, header_row), rows[header_row:]
    if suffix not in {".xlsx", ".xlsm"}:
        raise ValueError("Csak .xlsx, .xlsm vagy .csv fájl támogatott.")
    workbook = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    sheets = workbook.sheetnames
    selected = sheets[0]
    ws = workbook[selected]
    all_rows = list(ws.iter_rows(values_only=True))
    return sheets, selected, header_from_rows(all_rows, header_row), [list(r) for r in all_rows[header_row:]]


def header_from_rows(rows: list[Any], header_row: int) -> list[str]:
    index = max(header_row - 1, 0)
    if index >= len(rows):
        return []
    raw = list(rows[index])
    headers: list[str] = []
    seen: dict[str, int] = {}
    for i, value in enumerate(raw, start=1):
        label = clean_cell(value) or f"Oszlop {i}"
        if label in seen:
            seen[label] += 1
            label = f"{label} ({seen[label]})"
        else:
            seen[label] = 1
        headers.append(label)
    return headers


def normalise(value: str) -> str:
    repl = str.maketrans("áéíóöőúüűÁÉÍÓÖŐÚÜŰ", "aeiooouuuAEIOOOUUU")
    return re.sub(r"\s+", " ", value.translate(repl).lower()).strip()


def guess_mappings(headers: list[str]) -> dict[str, str]:
    norm_headers = {normalise(h): h for h in headers}
    guesses: dict[str, str] = {}
    for key, aliases in ALIASES.items():
        norm_aliases = [normalise(a) for a in aliases]
        for alias in norm_aliases:
            if alias in norm_headers:
                guesses[key] = norm_headers[alias]
                break
        if key not in guesses:
            for nh, original in norm_headers.items():
                if any(alias in nh for alias in norm_aliases):
                    guesses[key] = original
                    break
    return guesses


def row_is_empty(row: list[Any]) -> bool:
    return all(clean_cell(v) == "" for v in row)


def row_dict(headers: list[str], row: list[Any]) -> dict[str, str]:
    result: dict[str, str] = {}
    for i, header in enumerate(headers):
        result[header] = clean_cell(row[i] if i < len(row) else "")
    return result


def parse_date(value: str, field_name: str) -> str:
    raw = clean_cell(value)
    if not raw:
        raise ValueError(f"Hiányzik a dátum: {field_name}.")
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 8:
        return digits
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y%m%d")
        except ValueError:
            pass
    raise ValueError(f"Hibás dátumformátum ({field_name}): {raw}")


def value_for(row: dict[str, str], key: str, config: dict[str, Any]) -> str:
    header = (config.get("mapping") or {}).get(key) or ""
    defaults = config.get("defaults") or {}
    value = row.get(header, "") if header else ""
    return clean_cell(value) or clean_cell(defaults.get(key, ""))


def digits_only(value: str) -> str:
    return re.sub(r"\D", "", clean_cell(value))


def amount_to_field(value: str, decimals: str) -> str:
    raw = clean_cell(value).replace(" ", "").replace("\u00a0", "")
    if not raw:
        raise ValueError("Hiányzik az összeg.")
    raw = raw.replace(",", ".")
    try:
        decimal_amount = Decimal(raw)
        decimal_places = int(clean_cell(decimals) or "0")
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f"Hibás összeg vagy tizedesjegy: {value}") from exc
    multiplier = Decimal(10) ** decimal_places
    number = int((decimal_amount * multiplier).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    if number < 0:
        raise ValueError("Az összeg nem lehet negatív.")
    output = str(number).zfill(13)
    if len(output) > 13:
        raise ValueError(f"Az összeg túl hosszú a 13 karakteres mezőhöz: {value}")
    return output


def fit_text(value: str, length: int, pad: str = " ", left: bool = False) -> str:
    text = clean_cell(value)
    if len(text) > length:
        text = text[:length]
    if left:
        return text.rjust(length, pad)
    return text.ljust(length, pad)


def fit_digits(value: str, length: int, pad_left: bool = False) -> str:
    text = digits_only(value)
    if len(text) > length:
        text = text[:length]
    return text.zfill(length) if pad_left else text.ljust(length, " ")


def set_field(buffer: list[str], start: int, length: int, value: str) -> None:
    text = value
    if len(text) != length:
        raise ValueError(f"Belső mezőhossz hiba a {start}. pozíciónál: {len(text)} != {length}")
    begin = start - 1
    buffer[begin:begin + length] = list(text)


def normalise_status(value: str) -> str:
    raw = clean_cell(value).upper()
    if raw in {"", "NORMAL", "N", "0"}:
        return " "
    if raw in {"V", "VIBER"}:
        return "V"
    if raw in {"T", "TPLUSONE", "T+1"}:
        return "T"
    raise ValueError(f"Ismeretlen státusz: {value}")


def optional_date(value: str, field_name: str) -> str:
    if not clean_cell(value):
        return " " * 8
    return parse_date(value, field_name)


def build_payord_line(values: dict[str, str], sequence: int, config: dict[str, Any]) -> str:
    line = [" "] * 939
    decimals = clean_cell(values["decimals"] or "0")
    identifier = clean_cell(values["identifier"])
    if not identifier:
        identifier_date = parse_date(config.get("identifier_date") or date.today().strftime("%Y-%m-%d"), "azonosító dátuma")
        identifier = f"{identifier_date}{sequence:06d}"
    identifier = digits_only(identifier)
    if len(identifier) != 14:
        raise ValueError("Az azonosító 14 számjegy legyen, vagy hagyd üresen az automatikus generáláshoz.")

    set_field(line, 1, 6, "PAYORD")
    set_field(line, 7, 2, "DO")
    set_field(line, 9, 14, identifier)
    set_field(line, 23, 47, fit_digits(values["sender_account"], 47))
    set_field(line, 70, 1, "0")
    set_field(line, 71, 32, fit_text(values["sender_name"], 32))
    set_field(line, 220, 47, fit_digits(values["beneficiary_account"], 47))
    set_field(line, 267, 1, "0")
    set_field(line, 332, 32, fit_text(values["beneficiary_name"], 32))
    set_field(line, 472, 2, fit_text(values["beneficiary_country"], 2))
    set_field(line, 593, 96, fit_text(values["note"], 96))
    set_field(line, 689, 6, fit_text(values["document_no"], 6))
    set_field(line, 695, 4, fit_text(values["legal_code"], 4))
    set_field(line, 806, 3, fit_text(values["currency"].upper(), 3))
    set_field(line, 809, 1, fit_digits(decimals, 1))
    set_field(line, 810, 13, amount_to_field(values["amount"], decimals))
    set_field(line, 835, 8, parse_date(values["value_date"], "valutanap"))
    set_field(line, 843, 1, normalise_status(values["status"]))
    set_field(line, 938, 2, "00")
    return "".join(line)


def build_fx_payord_line(values: dict[str, str], sequence: int, config: dict[str, Any]) -> str:
    line = [" "] * 1046
    decimals = clean_cell(values["decimals"] or "2")
    identifier = clean_cell(values["identifier"])
    if not identifier:
        identifier_date = parse_date(config.get("identifier_date") or date.today().strftime("%Y-%m-%d"), "azonosító dátuma")
        identifier = f"{identifier_date}{sequence:06d}"
    identifier = digits_only(identifier)
    if len(identifier) != 14:
        raise ValueError("Az azonosító 14 számjegy legyen, vagy hagyd üresen az automatikus generáláshoz.")

    set_field(line, 1, 6, "PAYORD")
    set_field(line, 7, 2, "IN")
    set_field(line, 9, 14, identifier)
    set_field(line, 23, 47, fit_text(values["sender_account"], 47))
    set_field(line, 70, 1, fit_text(values["sender_account_type"] or "0", 1))
    set_field(line, 71, 140, fit_text(values["sender_name"], 140))
    set_field(line, 220, 47, fit_text(values["beneficiary_account"], 47))
    set_field(line, 267, 1, fit_text(values["beneficiary_account_type"] or "0", 1))
    set_field(line, 268, 64, fit_text(values["beneficiary_bank_name"], 64))
    set_field(line, 332, 140, fit_text(values["beneficiary_name"], 140))
    set_field(line, 472, 2, fit_text(values["beneficiary_country"], 2))
    set_field(line, 474, 14, fit_text(values["beneficiary_bank_swift"], 14))
    set_field(line, 489, 35, fit_text(values["correspondent_bank_1"], 35))
    set_field(line, 525, 1, fit_text(values["swift_copy"] or "N", 1))
    set_field(line, 526, 15, fit_text(values["fax_number"], 15))
    set_field(line, 541, 1, fit_text(values["custom_rate_use"] or "N", 1))
    set_field(line, 542, 15, fit_text(values["custom_rate"], 15))
    set_field(line, 558, 35, fit_text(values["bank_message"], 35))
    set_field(line, 593, 96, fit_text(values["note"], 96))
    set_field(line, 689, 4, fit_text(values["legal_code"], 4))
    set_field(line, 753, 24, fit_text(values["permit_no"], 24))
    set_field(line, 777, 8, optional_date(values["permit_date"], "deviza engedély dátuma"))
    set_field(line, 785, 1, fit_text(values["urgent_use"] or "N", 1))
    set_field(line, 786, 1, fit_text(values["urgent_execution"], 1))
    set_field(line, 788, 1, fit_text(values["fax_copy_execution"], 1))
    set_field(line, 789, 3, fit_text(values["payout_currency"].upper(), 3))
    set_field(line, 806, 3, fit_text(values["currency"].upper(), 3))
    set_field(line, 809, 1, fit_digits(decimals, 1))
    set_field(line, 810, 13, amount_to_field(values["amount"], decimals))
    set_field(line, 835, 8, optional_date(values["value_date"], "feldolgozás kezdő dátuma"))
    set_field(line, 855, 1, fit_text(values["cost_bearer"] or "1", 1))
    set_field(line, 938, 2, "00")
    set_field(line, 940, 2, fit_text(values["partner_bank_country"], 2))
    set_field(line, 942, 35, fit_text(values["correspondent_bank_2"], 35))
    set_field(line, 977, 35, fit_text(values["correspondent_bank_3"], 35))
    set_field(line, 1012, 35, fit_text(values["correspondent_bank_4"], 35))
    return "".join(line)


def get_format(format_id: str | None) -> dict[str, Any]:
    if format_id and format_id in FORMATS:
        return FORMATS[format_id]
    return FORMATS["erste_huf_payord"]


def fields_for_format(format_id: str | None) -> dict[str, dict[str, Any]]:
    return get_format(format_id).get("fields", FIELDS)


def validate_required(values: dict[str, str], row_number: int, fields: dict[str, dict[str, Any]]) -> None:
    missing = [info["label"] for key, info in fields.items() if info["required"] and not clean_cell(values.get(key, ""))]
    if missing:
        raise ValueError(f"{row_number}. sor: hiányzó kötelező mező: {', '.join(missing)}")


def convert_records(data: bytes, filename: str, config: dict[str, Any]) -> tuple[str, int]:
    header_row = 1
    format_id = config.get("format") or "erste_huf_payord"
    fields = fields_for_format(format_id)
    sheets, selected, headers, rows = workbook_rows(data, filename)
    if not headers:
        raise ValueError("Nem található fejlécsor.")
    output: list[str] = []
    sequence = 0
    errors: list[str] = []
    for offset, raw_row in enumerate(rows, start=header_row + 1):
        row = list(raw_row)
        if row_is_empty(row):
            continue
        data_row = row_dict(headers, row)
        values = {key: value_for(data_row, key, config) for key in fields}
        try:
            validate_required(values, offset, fields)
            sequence += 1
            if format_id == "erste_fx_payord":
                line = build_fx_payord_line(values, sequence, config)
            else:
                line = build_payord_line(values, sequence, config)
            output.append(line)
        except ValueError as exc:
            errors.append(f"{offset}. sor: {exc}")
            if len(errors) >= 12:
                break
    if errors:
        raise ValueError("\n".join(errors))
    if not output:
        raise ValueError("Nem találtam konvertálható adatsort.")
    return "\r\n".join(output) + "\r\n", len(output)


def build_template_xlsx(format_id: str | None = None) -> bytes:
    format_def = get_format(format_id)
    fields = fields_for_format(format_id)
    template_example = format_def.get("template_example") or TEMPLATE_EXAMPLE
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Utalasok"

    headers = [info["label"] for info in fields.values()]
    example = [template_example.get(key, "") for key in fields]
    sheet.append(headers)
    sheet.append(example)
    sheet.freeze_panes = "A2"

    header_fill = PatternFill("solid", fgColor="B51F2B")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D7DDE5")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    for cell in sheet[2]:
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        cell.border = border
    for index, header in enumerate(headers, start=1):
        width = max(14, min(32, len(header) + 4))
        sheet.column_dimensions[openpyxl.utils.get_column_letter(index)].width = width
    sheet.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(headers))}2"

    info_sheet = workbook.create_sheet("Mezo_info")
    info_sheet.append(["Mező", "Kötelező", "Megjegyzés"])
    for key, info in fields.items():
        note = info.get("help", "")
        if key == "status":
            note = "Üres = NORMAL, V = VIBER, T = TPLUSONE."
        elif key == "amount":
            note = "Számként vagy szövegként is megadható. A TXT-ben 13 karakterre, balról nullázva kerül."
        elif key == "value_date":
            note = "Dátumként vagy ÉÉÉÉHHNN formában add meg."
        elif key == "identifier":
            note = "Ha üres, a konverter automatikusan generálja."
        info_sheet.append([info["label"], "igen" if info["required"] else "nem", note])
    for cell in info_sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    for row in info_sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
    info_sheet.column_dimensions["A"].width = 28
    info_sheet.column_dimensions["B"].width = 12
    info_sheet.column_dimensions["C"].width = 78
    info_sheet.freeze_panes = "A2"

    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def default_settings() -> dict[str, Any]:
    return {
        "active_bank": "erste",
        "active_format": "erste_huf_payord",
        "formats": {},
    }


def load_settings_file() -> dict[str, Any]:
    if not SETTINGS_FILE.exists():
        return default_settings()
    try:
        data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        base = default_settings()
        base.update(data if isinstance(data, dict) else {})
        if not isinstance(base.get("formats"), dict):
            base["formats"] = {}
        return base
    except Exception:
        return default_settings()


def save_settings_file(data: dict[str, Any]) -> dict[str, Any]:
    cleaned = default_settings()
    cleaned.update(data if isinstance(data, dict) else {})
    if cleaned.get("active_bank") not in BANKS:
        cleaned["active_bank"] = "erste"
    if cleaned.get("active_format") not in FORMATS:
        cleaned["active_format"] = "erste_huf_payord"
    if not isinstance(cleaned.get("formats"), dict):
        cleaned["formats"] = {}
    for cfg in cleaned["formats"].values():
        if isinstance(cfg, dict):
            cfg.pop("sheet", None)
            cfg.pop("header_row", None)
    SETTINGS_FILE.write_text(json.dumps(cleaned, ensure_ascii=False, indent=2), encoding="utf-8")
    return cleaned


class Handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        if path in {"/", "/index.html"}:
            body = (
                HTML_PAGE
                .replace("__BANKS_JSON__", json.dumps(BANKS, ensure_ascii=False))
                .replace("__FORMATS_JSON__", json.dumps(FORMATS, ensure_ascii=False))
                .replace("__FIELDS_JSON__", json.dumps(FIELDS, ensure_ascii=False))
            )
            text_response(self, body)
            return
        if path == "/api/settings":
            json_response(self, load_settings_file())
            return
        if path == "/template.xlsx":
            params = parse_qs(parsed.query)
            format_id = (params.get("format") or [""])[0]
            data = build_template_xlsx(format_id)
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", "attachment; filename=payord_sablon.xlsx")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if path == "/template.csv":
            params = parse_qs(parsed.query)
            format_id = (params.get("format") or [""])[0]
            content = ";".join(info["label"] for info in fields_for_format(format_id).values()) + "\r\n"
            data = content.encode("utf-8-sig")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/csv; charset=utf-8")
            self.send_header("Content-Disposition", "attachment; filename=payord_sablon.csv")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        try:
            if self.path == "/api/settings":
                length = int(self.headers.get("Content-Length", "0") or "0")
                payload = self.rfile.read(length).decode("utf-8") if length else "{}"
                json_response(self, save_settings_file(json.loads(payload or "{}")))
                return
            if self.path == "/api/inspect":
                form = get_upload(self)
                filename, data = upload_bytes(form)
                sheets, selected, headers, rows = workbook_rows(data, filename)
                non_empty_rows = [r for r in rows if not row_is_empty(list(r))]
                sample = [[clean_cell(c) for c in list(r)[: len(headers)]] for r in non_empty_rows[:6]]
                json_response(self, {
                    "sheets": sheets,
                    "sheet": selected,
                    "headers": headers,
                    "guesses": guess_mappings(headers),
                    "sample": sample,
                    "data_rows": len(non_empty_rows),
                })
                return
            if self.path == "/api/convert":
                form = get_upload(self)
                filename, data = upload_bytes(form)
                config = json.loads(field_value(form, "config", "{}") or "{}")
                text, count = convert_records(data, filename, config)
                encoding = config.get("encoding") or "cp1250"
                codec = "cp1250" if encoding == "cp1250" else "utf-8"
                payload = text.encode(codec, errors="replace")
                self.send_response(HTTPStatus.OK)
                charset = "windows-1250" if codec == "cp1250" else "utf-8"
                self.send_header("Content-Type", f"text/plain; charset={charset}")
                self.send_header("Content-Disposition", "attachment; filename=payord_import.txt")
                self.send_header("X-Filename", "payord_import.txt")
                self.send_header("X-Record-Count", str(count))
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            self.send_error(HTTPStatus.NOT_FOUND)
        except Exception as exc:
            message = str(exc)
            json_response(self, {"error": message, "errors": [line for line in message.splitlines() if line]}, status=400)

    def log_message(self, fmt: str, *args: Any) -> None:
        print(f"{self.address_string()} - {fmt % args}")


def main() -> None:
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"PAYORD konverter fut: http://{HOST}:{PORT}")
    server.serve_forever()


if __name__ == "__main__":
    main()
