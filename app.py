from __future__ import annotations

import logging
import os

import csv
from email.parser import BytesParser
from email.policy import default as email_policy
import html
import io
import json
import re
import urllib.request
import uuid
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse
from xml.dom import minidom

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


HOST = "127.0.0.1"
PORT = 8765
SETTINGS_FILE = Path(__file__).with_name("settings.json")
ACCOUNTS_FILE = Path(__file__).with_name("own_accounts.json")
COMPANIES_FILE = Path(__file__).with_name("companies.json")
PARTNERS_FILE = Path(__file__).with_name("partners.json")
TRASH_FILE = Path(__file__).with_name("_trash.json")
AUDIT_FILE = Path(__file__).with_name("_audit.jsonl")

# --- Iteration 7: configurable runtime + schema versioning ---
SCHEMA_VERSION_FILE = Path(__file__).with_name("_schema_version.json")
CURRENT_SCHEMA_VERSION = 2
MNB_REFRESH_HOURS = float(os.environ.get("BANKI_MNB_REFRESH_HOURS", "24"))
AUDIT_MAX_LINES = int(os.environ.get("BANKI_AUDIT_MAX_LINES", "20000"))
LOG_LEVEL = os.environ.get("BANKI_LOG_LEVEL", "INFO").upper()
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL, logging.INFO),
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
log = logging.getLogger("banki")
BANK_REGISTRY_FILE = Path(__file__).with_name("bank_registry.json")

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

SEPA_FIELDS = {
    "identifier": {
        "label": "Azonosító",
        "required": False,
        "help": "Ha üres, automatikusan: dátum + sorszám.",
    },
    "sender_account": {"label": "Feladó számlaszám", "required": True},
    "sender_account_type": {"label": "Feladó számlaszám típusa", "required": False},
    "sender_name": {"label": "Feladó neve", "required": True},
    "beneficiary_account": {"label": "Címzett IBAN számlaszám", "required": True},
    "beneficiary_account_type": {"label": "Címzett számlaszám típusa", "required": False},
    "beneficiary_name": {"label": "Címzett neve", "required": True},
    "beneficiary_bank_swift": {"label": "Partner bank SWIFT kódja", "required": True},
    "note": {"label": "Közlemény", "required": False},
    "urgent_execution": {"label": "Sürgős teljesítés", "required": False},
    "payout_currency": {"label": "Kifizetendő deviza", "required": True},
    "currency": {"label": "Átutalandó deviza", "required": True},
    "decimals": {"label": "Tizedesjegyek száma", "required": True},
    "amount": {"label": "Átutalandó összeg", "required": True},
}

SEPA_TEMPLATE_EXAMPLE = {
    "identifier": "",
    "sender_account": "HU00111111112222222233333333",
    "sender_account_type": "0",
    "sender_name": "Minta Feladó Kft",
    "beneficiary_account": "AT451100000421840000",
    "beneficiary_account_type": "0",
    "beneficiary_name": "Minta Partner GmbH",
    "beneficiary_bank_swift": "BKAUATWW",
    "note": "Invoice payment",
    "urgent_execution": "",
    "payout_currency": "EUR",
    "currency": "EUR",
    "decimals": "2",
    "amount": "123.45",
}

KH_FX_FIELDS = {
    "identifier": {
        "label": "Azonosító",
        "required": False,
        "help": "Ha üres, automatikusan: dátum + sorszám.",
    },
    "sender_account": {"label": "Feladó számlaszám", "required": True},
    "sender_name": {"label": "Feladó neve", "required": True},
    "beneficiary_account": {"label": "Címzett számlaszám", "required": True},
    "beneficiary_bank_name": {"label": "Címzett bank neve", "required": True},
    "beneficiary_bank_address": {"label": "Címzett bank címe", "required": False},
    "beneficiary_name": {"label": "Címzett neve", "required": True},
    "beneficiary_bank_swift": {"label": "Partner bank SWIFT kódja", "required": True},
    "bank_message": {"label": "Közlemény a bank számára", "required": False},
    "note": {"label": "Közlemény", "required": False},
    "payout_currency": {"label": "Kifizetendő deviza", "required": True},
    "currency": {"label": "Átutalandó deviza", "required": True},
    "decimals": {"label": "Tizedesjegyek száma", "required": True},
    "amount": {"label": "Átutalandó összeg", "required": True},
    "value_date": {"label": "Valutanap", "required": False},
    "cost_bearer": {"label": "Költségviselés", "required": False},
    "urgent_execution": {"label": "Sürgős teljesítés", "required": False},
    "group_transfer": {"label": "Csoportos átutalás", "required": False},
}

KH_FX_TEMPLATE_EXAMPLE = {
    "identifier": "",
    "sender_account": "11111111-22222222-33333333",
    "sender_name": "Minta Feladó Kft",
    "beneficiary_account": "DE00123456781234567890",
    "beneficiary_bank_name": "Minta Bank",
    "beneficiary_bank_address": "Minta bankcím",
    "beneficiary_name": "Minta Partner GmbH",
    "beneficiary_bank_swift": "DEUTDEFF",
    "bank_message": "",
    "note": "Invoice payment",
    "payout_currency": "EUR",
    "currency": "EUR",
    "decimals": "2",
    "amount": "123.45",
    "value_date": "2026-06-17",
    "cost_bearer": "1",
    "urgent_execution": "N",
    "group_transfer": "N",
}

UNICREDIT_PAY_FIELDS = {
    "sender_account": {"label": "Terhelendő számlaszám", "required": True},
    "sender_currency": {"label": "Terhelendő számla devizája", "required": True},
    "beneficiary_account": {"label": "Kedvezményezett számlaszám", "required": True},
    "beneficiary_name": {"label": "Kedvezményezett neve", "required": True},
    "document_no": {"label": "Bizonylatszám", "required": False},
    "note": {"label": "Közlemény", "required": False},
    "amount": {"label": "Összeg", "required": True},
    "currency": {"label": "Összeg devizaneme", "required": True},
    "value_date": {"label": "Értéknap", "required": False},
    "legal_code": {"label": "Jogcímkód", "required": False},
    "beneficiary_country": {"label": "Jogosult országkódja", "required": True},
}

UNICREDIT_PAY_TEMPLATE_EXAMPLE = {
    "sender_account": "11111111-22222222-33333333",
    "sender_currency": "HUF",
    "beneficiary_account": "44444444-55555555-66666666",
    "beneficiary_name": "Minta Partner Kft",
    "document_no": "BIZ001",
    "note": "Számla kiegyenlítés",
    "amount": "12345",
    "currency": "HUF",
    "value_date": "2026-06-17",
    "legal_code": "",
    "beneficiary_country": "HU",
}

UNICREDIT_CCY_FIELDS = {
    "sender_account": {"label": "Terhelendő számlaszám", "required": True},
    "sender_currency": {"label": "Terhelendő számla devizája", "required": True},
    "beneficiary_bank_swift": {"label": "Kedvezményezett bank BIC kódja", "required": True},
    "beneficiary_bank_id": {"label": "Kedvezményezett bank azonosítója", "required": False},
    "beneficiary_bank_name": {"label": "Kedvezményezett bank neve", "required": True},
    "beneficiary_bank_address_1": {"label": "Kedvezményezett bank címe 1", "required": False},
    "beneficiary_bank_address_2": {"label": "Kedvezményezett bank címe 2", "required": False},
    "beneficiary_bank_address_3": {"label": "Kedvezményezett bank címe 3", "required": False},
    "correspondent_bank_swift": {"label": "Levelező bank BIC kódja", "required": False},
    "correspondent_bank_id": {"label": "Levelező bank azonosítója", "required": False},
    "correspondent_bank_name": {"label": "Levelező bank neve", "required": False},
    "correspondent_bank_address_1": {"label": "Levelező bank címe 1", "required": False},
    "correspondent_bank_address_2": {"label": "Levelező bank címe 2", "required": False},
    "correspondent_bank_address_3": {"label": "Levelező bank címe 3", "required": False},
    "beneficiary_account": {"label": "Kedvezményezett számlaszáma", "required": True},
    "beneficiary_name": {"label": "Kedvezményezett neve", "required": True},
    "beneficiary_address_1": {"label": "Kedvezményezett címe 1", "required": False},
    "beneficiary_address_2": {"label": "Kedvezményezett címe 2", "required": False},
    "beneficiary_address_3": {"label": "Kedvezményezett címe 3", "required": False},
    "note": {"label": "Közlemény", "required": False},
    "cost_bearer": {"label": "Banki költségek megosztása", "required": True},
    "payout_currency": {"label": "Teljesítés devizaneme", "required": True},
    "amount": {"label": "Összeg", "required": True},
    "currency": {"label": "Összeg devizaneme", "required": True},
    "value_date": {"label": "Értéknap", "required": False},
    "urgent_execution": {"label": "Sürgősség átutalás flag", "required": False},
    "hold_flag": {"label": "HOLD flag", "required": False},
    "chqb_flag": {"label": "CHQB flag", "required": False},
    "deal_ticket_flag": {"label": "Deal Ticket flag", "required": False},
    "deal_ticket_date": {"label": "Deal Ticket dátuma", "required": False},
    "deal_ticket_sequence": {"label": "Deal Ticket sorszáma", "required": False},
    "beneficiary_country": {"label": "Kedvezményezett országkódja", "required": False},
    "legal_code": {"label": "Statisztikai jogcímkód", "required": True},
}

UNICREDIT_CCY_TEMPLATE_EXAMPLE = {
    "sender_account": "11111111-22222222-33333333",
    "sender_currency": "HUF",
    "beneficiary_bank_swift": "DEUTDEFF",
    "beneficiary_bank_id": "",
    "beneficiary_bank_name": "Minta Bank",
    "beneficiary_bank_address_1": "Minta cím 1",
    "beneficiary_bank_address_2": "",
    "beneficiary_bank_address_3": "",
    "correspondent_bank_swift": "",
    "correspondent_bank_id": "",
    "correspondent_bank_name": "",
    "correspondent_bank_address_1": "",
    "correspondent_bank_address_2": "",
    "correspondent_bank_address_3": "",
    "beneficiary_account": "DE00123456781234567890",
    "beneficiary_name": "Minta Partner GmbH",
    "beneficiary_address_1": "Minta cím 1",
    "beneficiary_address_2": "",
    "beneficiary_address_3": "",
    "note": "Invoice payment",
    "cost_bearer": "SHA",
    "payout_currency": "EUR",
    "amount": "123.45",
    "currency": "EUR",
    "value_date": "2026-06-17",
    "urgent_execution": "N",
    "hold_flag": "N",
    "chqb_flag": "N",
    "deal_ticket_flag": "N",
    "deal_ticket_date": "",
    "deal_ticket_sequence": "",
    "beneficiary_country": "DE",
    "legal_code": "000",
}

OTP_HUF_FIELDS = {
    "identifier": {"label": "Azonosító", "required": False, "help": "Ha üres, automatikusan: dátum + sorszám."},
    "sender_account": {"label": "Feladó számlaszám", "required": True},
    "sender_name": {"label": "Feladó neve", "required": True},
    "beneficiary_account": {"label": "Címzett számlaszám", "required": True},
    "beneficiary_name": {"label": "Címzett neve", "required": True},
    "note": {"label": "Közlemény", "required": False},
    "document_no": {"label": "Bizonylatszám", "required": False},
    "currency": {"label": "Deviza", "required": True},
    "decimals": {"label": "Tizedesjegyek száma", "required": True},
    "amount": {"label": "Átutalandó összeg", "required": True},
    "value_date": {"label": "Értéknap", "required": False},
    "process_mode": {"label": "Feldolgozási mód", "required": False},
}

OTP_HUF_TEMPLATE_EXAMPLE = {
    "identifier": "",
    "sender_account": "11111111-22222222-33333333",
    "sender_name": "Minta Feladó Kft",
    "beneficiary_account": "44444444-55555555-66666666",
    "beneficiary_name": "Minta Partner Kft",
    "note": "Számla kiegyenlítés",
    "document_no": "BIZ001",
    "currency": "HUF",
    "decimals": "0",
    "amount": "12345",
    "value_date": "",
    "process_mode": "",
}

OTP_FX_FIELDS = {
    "identifier": {"label": "Azonosító", "required": False, "help": "Ha üres, automatikusan: dátum + sorszám."},
    "sender_account": {"label": "Feladó számlaszám", "required": True},
    "sender_account_type": {"label": "Feladó számlaszám típusa", "required": True},
    "sender_name": {"label": "Feladó neve", "required": True},
    "beneficiary_account": {"label": "Címzett számlaszám", "required": True},
    "beneficiary_account_type": {"label": "Címzett számlaszám típusa", "required": True},
    "beneficiary_bank_name": {"label": "Címzett bank neve", "required": True},
    "beneficiary_bank_address": {"label": "Címzett bank címe", "required": False},
    "beneficiary_name": {"label": "Címzett neve", "required": True},
    "beneficiary_bank_swift": {"label": "Partner bank azonosítója", "required": True},
    "correspondent_bank_1": {"label": "Levelező bank", "required": False},
    "bank_message": {"label": "Megjegyzés a bank számára", "required": False},
    "note": {"label": "Közlemény", "required": False},
    "permit_no": {"label": "Deviza engedély száma", "required": False},
    "permit_date": {"label": "Deviza engedély dátuma", "required": False},
    "urgent_execution": {"label": "Sürgős teljesítés", "required": False},
    "payout_currency": {"label": "Kifizetendő deviza", "required": True},
    "currency": {"label": "Átutalandó deviza", "required": True},
    "decimals": {"label": "Tizedesjegyek száma", "required": True},
    "amount": {"label": "Átutalandó összeg", "required": True},
    "cost_bearer": {"label": "Költségviselés", "required": True},
}

OTP_FX_TEMPLATE_EXAMPLE = {
    "identifier": "",
    "sender_account": "HU00111111112222222233333333",
    "sender_account_type": "9",
    "sender_name": "Minta Feladó Kft",
    "beneficiary_account": "DE00123456781234567890",
    "beneficiary_account_type": "9",
    "beneficiary_bank_name": "Minta Bank",
    "beneficiary_bank_address": "Minta bankcím",
    "beneficiary_name": "Minta Partner GmbH",
    "beneficiary_bank_swift": "DEUTDEFF",
    "correspondent_bank_1": "",
    "bank_message": "",
    "note": "Invoice payment",
    "permit_no": "",
    "permit_date": "",
    "urgent_execution": "1",
    "payout_currency": "EUR",
    "currency": "EUR",
    "decimals": "2",
    "amount": "123.45",
    "cost_bearer": "1",
}

RAIFFEISEN_HUF_FIELDS = {
    "amount": {"label": "Összeg", "required": True},
    "value_date": {"label": "Értéknap", "required": False},
    "sender_account": {"label": "Kezdeményező pénzforgalmi jelzőszáma", "required": True},
    "beneficiary_account": {"label": "Kedvezményezett pénzforgalmi jelzőszáma", "required": True},
    "beneficiary_name": {"label": "Kedvezményezett neve", "required": True},
    "beneficiary_country": {"label": "Kedvezményezett országkódja", "required": False},
    "legal_code": {"label": "Jogcím", "required": False},
    "note": {"label": "Közlemény", "required": False},
    "external_ref": {"label": "Külső referencia", "required": False},
    "document_no": {"label": "Bizonylatszám", "required": False},
    "internal_note": {"label": "Belső megjegyzés", "required": False},
}

RAIFFEISEN_HUF_TEMPLATE_EXAMPLE = {
    "amount": "12345",
    "value_date": "2026-06-17",
    "sender_account": "11111111-22222222-33333333",
    "beneficiary_account": "44444444-55555555-66666666",
    "beneficiary_name": "Minta Partner Kft",
    "beneficiary_country": "HU",
    "legal_code": "",
    "note": "Számla kiegyenlítés",
    "external_ref": "",
    "document_no": "BIZ001",
    "internal_note": "",
}

RAIFFEISEN_FX_FIELDS = {
    "currency": {"label": "Átutalandó összeg devizaneme", "required": True},
    "sender_currency": {"label": "Megbízó számla devizaneme", "required": False},
    "amount": {"label": "Összeg", "required": True},
    "value_date": {"label": "Értéknap", "required": False},
    "sender_account": {"label": "Terhelendő számlaszám", "required": True},
    "beneficiary_name": {"label": "Kedvezményezett neve", "required": True},
    "beneficiary_bank_name": {"label": "Kedvezményezett bank neve", "required": True},
    "beneficiary_bank_country": {"label": "Kedvezményezett bank ország", "required": False},
    "beneficiary_bank_address_1": {"label": "Kedvezményezett bank címe 1", "required": False},
    "beneficiary_bank_address_2": {"label": "Kedvezményezett bank címe 2", "required": False},
    "beneficiary_account": {"label": "Kedvezményezett számlaszáma", "required": True},
    "legal_code": {"label": "Jogcímkód", "required": False},
    "note": {"label": "Közlemény", "required": False},
    "commission_bearer": {"label": "Ügyfél jutalék fizetője", "required": True},
    "other_fee_bearer": {"label": "Egyéb jutalék fizetője", "required": True},
    "permit_no": {"label": "Engedély szám", "required": False},
    "document_no": {"label": "Bizonylatszám", "required": False},
    "beneficiary_bank_swift": {"label": "Kedvezményezett swift címe", "required": False},
    "amount_currency_mode": {"label": "Összeg devizaneme jel", "required": False},
    "payment_method": {"label": "Teljesítés módja", "required": False},
    "priority": {"label": "Prioritás", "required": False},
    "item_type": {"label": "Tétel típusa", "required": False},
    "iban_flag": {"label": "IBAN jelzés", "required": False},
    "beneficiary_country": {"label": "Kedvezményezett országkódja", "required": False},
    "external_ref": {"label": "Külső referencia", "required": False},
}

RAIFFEISEN_FX_TEMPLATE_EXAMPLE = {
    "currency": "EUR",
    "sender_currency": "HUF",
    "amount": "123.45",
    "value_date": "2026-06-17",
    "sender_account": "11111111-22222222-33333333",
    "beneficiary_name": "Minta Partner GmbH",
    "beneficiary_bank_name": "Minta Bank",
    "beneficiary_bank_country": "DE",
    "beneficiary_bank_address_1": "Minta cím 1",
    "beneficiary_bank_address_2": "",
    "beneficiary_account": "DE00123456781234567890",
    "legal_code": "",
    "note": "Invoice payment",
    "commission_bearer": "0",
    "other_fee_bearer": "0",
    "permit_no": "",
    "document_no": "BIZ001",
    "beneficiary_bank_swift": "DEUTDEFF",
    "amount_currency_mode": " ",
    "payment_method": " ",
    "priority": " ",
    "item_type": " ",
    "iban_flag": "1",
    "beneficiary_country": "DE",
    "external_ref": "",
}

PAIN_FIELDS = {
    "message_id": {"label": "Üzenet azonosító", "required": False},
    "payment_id": {"label": "Fizetési blokk azonosító", "required": False},
    "debtor_name": {"label": "Terhelendő fél neve", "required": True},
    "sender_account": {"label": "Terhelendő számla", "required": True},
    "sender_currency": {"label": "Terhelendő számla devizája", "required": False},
    "beneficiary_name": {"label": "Kedvezményezett neve", "required": True},
    "beneficiary_account": {"label": "Kedvezményezett számlaszáma / IBAN", "required": True},
    "beneficiary_bank_swift": {"label": "Kedvezményezett bank BIC/SWIFT", "required": False},
    "beneficiary_country": {"label": "Kedvezményezett országkódja", "required": False},
    "amount": {"label": "Összeg", "required": True},
    "currency": {"label": "Deviza", "required": True},
    "value_date": {"label": "Értéknap", "required": True},
    "end_to_end_id": {"label": "EndToEnd azonosító", "required": False},
    "cost_bearer": {"label": "Költségviselés", "required": False},
    "note": {"label": "Közlemény", "required": False},
}

PAIN_TEMPLATE_EXAMPLE = {
    "message_id": "",
    "payment_id": "",
    "debtor_name": "Minta Feladó Kft",
    "sender_account": "HU00111111112222222233333333",
    "sender_currency": "EUR",
    "beneficiary_name": "Minta Partner GmbH",
    "beneficiary_account": "AT451100000421840000",
    "beneficiary_bank_swift": "BKAUATWW",
    "beneficiary_country": "AT",
    "amount": "123.45",
    "currency": "EUR",
    "value_date": "2026-06-17",
    "end_to_end_id": "",
    "cost_bearer": "SLEV",
    "note": "Invoice payment",
}

BANKS = {
    "erste": {"label": "Erste Bank"},
    "kh": {"label": "K&H Bank"},
    "mbh": {"label": "MBH Bank"},
    "unicredit": {"label": "UniCredit Bank"},
    "raiffeisen": {"label": "Raiffeisen Bank"},
    "otp": {"label": "OTP Bank"},
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
    "erste_sepa_payord": {
        "bank": "erste",
        "label": "SEPA átutalás - PAYORD (IN / SCE)",
        "short_label": "PAYORD / IN SEPA",
        "badge": "PAYORD / IN · SEPA .SCE · 1048 karakter soronként CR/LF-fel",
        "description": "Erste SEPA (.SCE) átutalási import: a deviza PAYORD (IN) formátum szűkített mezőkészlete, EUR SEPA utaláshoz.",
        "default_encoding": "cp1250",
        "line_length": 1048,
        "fields": SEPA_FIELDS,
        "template_example": SEPA_TEMPLATE_EXAMPLE,
    },
    "kh_huf_payord": {
        "bank": "kh",
        "label": "Forint átutalás - PAYORD (DO)",
        "short_label": "PAYORD / DO",
        "badge": "K&H PAYORD / DO · 941 karakter soronként CR/LF-fel",
        "description": "K&H Electra forint átutalási import, EDIFACT PAYORD DO rekord.",
        "default_encoding": "cp1250",
        "line_length": 941,
        "fields": FIELDS,
        "template_example": TEMPLATE_EXAMPLE,
    },
    "kh_fx_payord": {
        "bank": "kh",
        "label": "Deviza átutalás - PAYORD (IN)",
        "short_label": "PAYORD / IN",
        "badge": "K&H PAYORD / IN · 941 karakter soronként CR/LF-fel",
        "description": "K&H Electra deviza átutalási import, bank-specifikus 939 adatkarakter + CR/LF rekorddal.",
        "default_encoding": "cp1250",
        "line_length": 941,
        "fields": KH_FX_FIELDS,
        "template_example": KH_FX_TEMPLATE_EXAMPLE,
    },
    "unicredit_huf_pay": {
        "bank": "unicredit",
        "label": "Forint átutalás - PAY",
        "short_label": "PAY",
        "badge": "UniCredit PAY · 256 byte rekordok · header/data/trailer",
        "description": "UniCredit SpectraNet forintátutalási import: 256 byte-os PAY header, data és trailer rekordokkal.",
        "default_encoding": "cp1250",
        "line_length": 256,
        "fields": UNICREDIT_PAY_FIELDS,
        "template_example": UNICREDIT_PAY_TEMPLATE_EXAMPLE,
    },
    "unicredit_fx_ccy": {
        "bank": "unicredit",
        "label": "Deviza átutalás - CCY",
        "short_label": "CCY",
        "badge": "UniCredit CCY · 800 byte rekordok · header/data/trailer",
        "description": "UniCredit SpectraNet devizaátutalási import: 800 byte-os CCY header, data és trailer rekordokkal.",
        "default_encoding": "cp1250",
        "line_length": 800,
        "fields": UNICREDIT_CCY_FIELDS,
        "template_example": UNICREDIT_CCY_TEMPLATE_EXAMPLE,
    },
    "kh_sepa_pain": {
        "bank": "kh",
        "label": "SEPA átutalás - PAIN XML",
        "short_label": "PAIN XML SEPA",
        "badge": "K&H SEPA · PAIN XML",
        "description": "K&H Electra SEPA import PAIN XML formátumban.",
        "default_encoding": "utf-8",
        "line_length": None,
        "fields": PAIN_FIELDS,
        "template_example": PAIN_TEMPLATE_EXAMPLE,
        "builder": "pain_sepa",
    },
    "mbh_huf": {
        "bank": "mbh",
        "label": "HUF átutalás - PAIN XML",
        "short_label": "MBH HUF",
        "badge": "MBH HUF · ISO20022 PAIN XML",
        "description": "MBH Vállalati Netbank / Direct Bank ISO20022 pain.001 import HUF utaláshoz.",
        "default_encoding": "utf-8",
        "line_length": None,
        "fields": PAIN_FIELDS,
        "template_example": {**PAIN_TEMPLATE_EXAMPLE, "currency": "HUF", "sender_currency": "HUF", "beneficiary_account": "44444444-55555555-66666666", "beneficiary_country": "HU", "cost_bearer": "SHAR"},
        "builder": "pain_huf",
    },
    "mbh_fx": {
        "bank": "mbh",
        "label": "Deviza átutalás - PAIN XML",
        "short_label": "MBH Deviza",
        "badge": "MBH Deviza · ISO20022 PAIN XML",
        "description": "MBH Vállalati Netbank / Direct Bank ISO20022 pain.001 import deviza utaláshoz.",
        "default_encoding": "utf-8",
        "line_length": None,
        "fields": PAIN_FIELDS,
        "template_example": PAIN_TEMPLATE_EXAMPLE,
        "builder": "pain_fx",
    },
    "mbh_sepa": {
        "bank": "mbh",
        "label": "SEPA átutalás - PAIN XML",
        "short_label": "MBH SEPA",
        "badge": "MBH SEPA · ISO20022 PAIN XML",
        "description": "MBH Vállalati Netbank / Direct Bank ISO20022 pain.001 import SEPA utaláshoz.",
        "default_encoding": "utf-8",
        "line_length": None,
        "fields": PAIN_FIELDS,
        "template_example": PAIN_TEMPLATE_EXAMPLE,
        "builder": "pain_sepa",
    },
    "otp_huf": {
        "bank": "otp",
        "label": "HUF átutalás - PAYORD (DO)",
        "short_label": "PAYORD / DO",
        "badge": "OTP PAYORD / DO · 941 karakter soronként CR/LF-fel",
        "description": "OTPdirekt Electra forint átutalási EDIFACT PAYORD import.",
        "default_encoding": "cp1250",
        "line_length": 941,
        "fields": OTP_HUF_FIELDS,
        "template_example": OTP_HUF_TEMPLATE_EXAMPLE,
    },
    "otp_fx": {
        "bank": "otp",
        "label": "Deviza átutalás - PAYORD (IN)",
        "short_label": "PAYORD / IN",
        "badge": "OTP PAYORD / IN · 941 karakter soronként CR/LF-fel",
        "description": "OTPdirekt Electra bankon kívüli deviza átutalási EDIFACT PAYORD import.",
        "default_encoding": "cp1250",
        "line_length": 941,
        "fields": OTP_FX_FIELDS,
        "template_example": OTP_FX_TEMPLATE_EXAMPLE,
    },
    "raiffeisen_huf": {
        "bank": "raiffeisen",
        "label": "HUF átutalás - Raiffeisen DBF",
        "short_label": "Raiffeisen DBF HUF",
        "badge": "Raiffeisen DBF HUF · 251 karakter soronként CR/LF-fel",
        "description": "Raiffeisen Electra forint átutalási DBF import.",
        "default_encoding": "cp852",
        "line_length": 251,
        "fields": RAIFFEISEN_HUF_FIELDS,
        "template_example": RAIFFEISEN_HUF_TEMPLATE_EXAMPLE,
    },
    "raiffeisen_fx": {
        "bank": "raiffeisen",
        "label": "Deviza átutalás - Raiffeisen DBF",
        "short_label": "Raiffeisen DBF Deviza",
        "badge": "Raiffeisen DBF Deviza · 495 karakter soronként CR/LF-fel",
        "description": "Raiffeisen Electra deviza átutalási DBF import.",
        "default_encoding": "cp852",
        "line_length": 495,
        "fields": RAIFFEISEN_FX_FIELDS,
        "template_example": RAIFFEISEN_FX_TEMPLATE_EXAMPLE,
    },
    "raiffeisen_sepa": {
        "bank": "raiffeisen",
        "label": "SEPA átutalás - Raiffeisen DBF",
        "short_label": "Raiffeisen DBF SEPA",
        "badge": "Raiffeisen DBF SEPA CT · 495 karakter soronként CR/LF-fel",
        "description": "Raiffeisen Electra SEPA CT import a deviza DBF formátum FTMOD=2 jelölésével.",
        "default_encoding": "cp852",
        "line_length": 495,
        "fields": RAIFFEISEN_FX_FIELDS,
        "template_example": {**RAIFFEISEN_FX_TEMPLATE_EXAMPLE, "currency": "EUR", "payment_method": "2", "iban_flag": "1"},
    },
}

CURRENCIES = [
    "HUF", "EUR", "AUD", "BGN", "BRL", "CAD", "CHF", "CNY", "CZK", "DKK", "GBP", "HKD",
    "HRK", "IDR", "ILS", "INR", "ISK", "JPY", "KRW", "MXN", "MYR", "NOK", "NZD", "PHP",
    "PLN", "RON", "RSD", "RUB", "SEK", "SGD", "THB", "TRY", "UAH", "USD", "ZAR", "ATS",
    "AUP", "BEF", "BGL", "CSD", "CSK", "DDM", "DEM", "EEK", "EGP", "ESP", "FIM", "FRF",
    "GHP", "GRD", "IEP", "ITL", "KPW", "KWD", "LBP", "LTL", "LUF", "LVL", "MNT", "NLG",
    "OAL", "OBL", "OFR", "ORB", "PKR", "PTE", "ROL", "SDP", "SIT", "SKK", "SUR", "VND",
    "XEU", "XTR", "YUD",
]

ALIASES = {
    "identifier": ["azonosító", "azonosito", "id", "megbízás azonosító", "megbizas azonosito"],
    "sender_account": ["feladó számlaszám", "felado szamlaszam", "forrás számla", "forras szamla"],
    "sender_currency": ["terhelendő számla devizája", "terhelendo szamla devizaja", "feladó deviza", "felado deviza"],
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
    "beneficiary_bank_id": ["kedvezményezett bank azonosítója", "kedvezmenyezett bank azonositoja", "bank azonosító"],
    "beneficiary_bank_address": ["címzett bank címe", "cimzett bank cime", "kedvezményezett bank címe"],
    "beneficiary_bank_address_1": ["kedvezményezett bank címe 1", "kedvezmenyezett bank cime 1", "címzett bank címe 1"],
    "beneficiary_bank_address_2": ["kedvezményezett bank címe 2", "kedvezmenyezett bank cime 2", "címzett bank címe 2"],
    "beneficiary_bank_address_3": ["kedvezményezett bank címe 3", "kedvezmenyezett bank cime 3", "címzett bank címe 3"],
    "beneficiary_name": [
        "címzett neve",
        "cimzett neve",
        "kedvezményezett neve",
        "kedvezmenyezett neve",
        "partner neve",
    ],
    "beneficiary_country": ["ország", "orszag", "címzett országa", "cimzett orszaga"],
    "beneficiary_bank_swift": ["swift", "swift kód", "swift kod", "partner bank swift", "partner bank swift kódja"],
    "correspondent_bank_swift": ["levelező bank bic", "levelezo bank bic", "levelező bank swift"],
    "correspondent_bank_id": ["levelező bank azonosítója", "levelezo bank azonositoja"],
    "correspondent_bank_name": ["levelező bank neve", "levelezo bank neve"],
    "beneficiary_address_1": ["kedvezményezett címe 1", "kedvezmenyezett cime 1", "címzett címe 1"],
    "beneficiary_address_2": ["kedvezményezett címe 2", "kedvezmenyezett cime 2", "címzett címe 2"],
    "beneficiary_address_3": ["kedvezményezett címe 3", "kedvezmenyezett cime 3", "címzett címe 3"],
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
    "currency": ["deviza", "currency", "pénznem", "penznem"],
    "decimals": ["tizedesjegyek száma", "tizedesjegy", "decimals"],
    "amount": ["összeg", "osszeg", "átutalandó összeg", "atutalando osszeg", "amount"],
    "value_date": ["valutanap", "értéknap", "erteknap", "dátum", "datum"],
    "cost_bearer": ["költségviselés", "koltsegviseles"],
    "partner_bank_country": ["partner bank országkódja", "partner bank orszagkodja"],
    "correspondent_bank_2": ["levelező bank 2", "levelezo bank 2", "levelező bank 2. sor"],
    "correspondent_bank_3": ["levelező bank 3", "levelezo bank 3", "levelező bank 3. sor"],
    "correspondent_bank_4": ["levelező bank 4", "levelezo bank 4", "levelező bank 4. sor"],
    "correspondent_bank_address_1": ["levelező bank címe 1", "levelezo bank cime 1"],
    "correspondent_bank_address_2": ["levelező bank címe 2", "levelezo bank cime 2"],
    "correspondent_bank_address_3": ["levelező bank címe 3", "levelezo bank cime 3"],
    "hold_flag": ["hold flag", "hold"],
    "chqb_flag": ["chqb flag", "chqb"],
    "deal_ticket_flag": ["deal ticket flag"],
    "deal_ticket_date": ["deal ticket dátuma", "deal ticket datuma"],
    "deal_ticket_sequence": ["deal ticket sorszáma", "deal ticket sorszama"],
    "status": ["státusz", "statusz", "status"],
    "external_ref": ["külső referencia", "kulso referencia", "kulsref", "external ref"],
    "internal_note": ["belső megjegyzés", "belso megjegyzes"],
    "beneficiary_bank_country": ["kedvezményezett bank ország", "kedvezmenyezett bank orszag", "bank ország"],
    "process_mode": ["feldolgozási mód", "feldolgozasi mod", "process mode"],
    "commission_bearer": ["ügyfél jutalék fizetője", "ugyfel jutalek fizetoje", "fbkmk"],
    "other_fee_bearer": ["egyéb jutalék fizetője", "egyeb jutalek fizetoje", "fekmk"],
    "amount_currency_mode": ["összeg devizaneme jel", "osszeg devizaneme jel", "fodev"],
    "payment_method": ["teljesítés módja", "teljesites modja", "ftmod"],
    "priority": ["prioritás", "prioritas", "fprior"],
    "item_type": ["tétel típusa", "tetel tipusa", "fttip"],
    "iban_flag": ["iban jelzés", "iban jelzes", "fiban"],
    "message_id": ["üzenet azonosító", "uzenet azonosito", "message id", "msgid"],
    "payment_id": ["fizetési blokk azonosító", "fizetesi blokk azonosito", "payment id", "pmtinfid"],
    "debtor_name": ["terhelendő fél neve", "terhelendo fel neve", "adós neve", "ados neve", "debtor name"],
    "end_to_end_id": ["endtoend azonosító", "endtoend azonosito", "end to end id"],
}

HTML_PAGE = r"""<!doctype html>
<html lang="hu">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Banki import konvertáló</title>
  <script>
    try {
      if (localStorage.getItem("banki.theme") === "dark") {
        document.documentElement.classList.add("theme-dark");
      }
      document.documentElement.dataset.lang = localStorage.getItem("banki.lang") || "hu";
    } catch (_) {}
  </script>
  <link rel="preconnect" href="https://rsms.me/">
  <link rel="stylesheet" href="https://rsms.me/inter/inter.css">
  <link rel="stylesheet" href="/static/styles-base.css">
  <link rel="stylesheet" href="/static/tokens.css">
  <link rel="manifest" href="/manifest.webmanifest">
  <meta name="theme-color" content="#c0263b">
  <link rel="apple-touch-icon" href="/static/icon-192.png">
</head>
<body>
  <div class="shell">
    <header>
      <div>
        <h1>Banki import konvertáló</h1>
      </div>
      <div class="header-meta">
        <div id="registryPill" class="status-pill warn" role="status" aria-live="polite">MNB tábla: betöltés</div>
        <div id="formatBadge" class="badge">PAYORD / DO · 941 karakter soronként CR/LF-fel</div>
      </div>
    </header>

    <section class="commandbar">
      <div class="command-summary sr-only">
        <strong id="commandFormatName">PAYORD / DO</strong>
        <span id="statusBox" role="status" aria-live="polite">Nyisd meg az Import menüt, válassz fájlt, majd olvasd be.</span>
      </div>
      <div class="command-actions">
        <select id="companySelect" title="Aktív cég"></select>
        <button id="companiesBtn" class="secondary" type="button">Cégek</button>
        <button id="openImportBtn" class="secondary" type="button">Import</button>
        <button id="accountsBtn" class="secondary" type="button">Saját bankszámlák</button>
        <button id="partnersBtn" class="secondary" type="button">Partnerek</button>
        <button id="convertBtn" class="primary" disabled aria-describedby="convertDisabledReason">TXT letöltése</button>
        <span id="convertDisabledReason" class="sr-only">Tölts fel fájlt és olvasd be az importot.</span>
        <a id="templateLink" class="button-link ghost" href="/template.xlsx" download>Excel sablon</a>
        <a id="backupLink" class="button-link ghost" href="/api/backup" download title="Összes JSON + audit log + séma verzió ZIP-ben">Mentés (ZIP)</a>
        <button id="helpBtn" class="ghost" type="button">? Súgó</button>
      </div>
    </section>

    <div class="layout">
      <main class="panel">
        <h2>Beolvasás eredménye</h2>
        <div class="panel-body">
          <div id="resultSummary" class="result-grid" aria-live="polite">
            <div class="metric"><strong>-</strong><span>Fejléc</span></div>
            <div class="metric"><strong>-</strong><span>Adatsor</span></div>
            <div class="metric"><strong>-</strong><span>Formátum</span></div>
          </div>
          <div id="errorArea" class="error-list" role="alert" aria-live="assertive" style="margin-top:14px;"></div>
          <section class="empty-state" id="resultEmptyState" style="margin-top:16px;">
            <h3>Még nincs beolvasott importfájl</h3>
            <p>Indítsd el az importot három lépésben.</p>
            <ol>
              <li>Válaszd ki a cégedet a fejlécben.</li>
              <li>Nyisd meg az <strong>Import</strong> panelt, válassz bankot, formátumot és fájlt.</li>
              <li>Beolvasás után töltsd le a <strong>TXT</strong> importfájlt.</li>
            </ol>
            <div class="cta-row">
              <button class="primary" type="button" onclick="document.getElementById('openImportBtn').click()">Import indítása</button>
              <a class="button-link ghost" href="/template.xlsx" download>Excel sablon letöltése</a>
            </div>
          </section>
          <div style="margin-top:16px;">
            <h3 style="font-size:13px;margin:0 0 8px;color:var(--ink-500);text-transform:uppercase;letter-spacing:.05em;">Beolvasott minta</h3>
            <div id="sampleArea" class="sample-wrap"></div>
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

  <import-dialog id="importDialog"></import-dialog>
  <accounts-dialog id="accountsDialog"></accounts-dialog>
  <companies-dialog id="companiesDialog"></companies-dialog>
  <company-edit-dialog id="companyEditDialog"></company-edit-dialog>
  <partners-dialog id="partnersDialog"></partners-dialog>
  <account-edit-dialog id="accountEditDialog"></account-edit-dialog>
  <partner-edit-dialog id="partnerEditDialog"></partner-edit-dialog>

  <help-dialog id="helpDialog"></help-dialog>

<script>
window.__BANKS = __BANKS_JSON__;
window.__FORMATS = __FORMATS_JSON__;
window.__FIELDS = __FIELDS_JSON__;
window.__CURRENCIES = __CURRENCIES_JSON__;
</script>
<script defer src="/static/js/core-dom.js"></script>
<script defer src="/static/js/components/dialog-base.js"></script>
<script defer src="/static/js/components/help-dialog.js"></script>
<script defer src="/static/js/components/partner-edit-dialog.js"></script>
<script defer src="/static/js/components/import-dialog.js"></script>
<script defer src="/static/js/components/accounts-dialog.js"></script>
<script defer src="/static/js/components/companies-dialog.js"></script>
<script defer src="/static/js/components/company-edit-dialog.js"></script>
<script defer src="/static/js/components/partners-dialog.js"></script>
<script defer src="/static/js/components/account-edit-dialog.js"></script>
<script defer src="/static/js/companies.js"></script>
<script defer src="/static/js/samples.js"></script>
<script defer src="/static/js/accounts.js"></script>
<script defer src="/static/js/partners.js"></script>
<script defer src="/static/js/registry.js"></script>
<script defer src="/static/js/convert.js"></script>
<script defer src="/static/js/dialogs.js"></script>
<script defer src="/static/app.js"></script>
<script defer src="/static/js/toast.js"></script>
<script defer src="/static/js/theme.js"></script>
<script defer src="/static/js/preview.js"></script>
<script defer src="/static/js/filter.js"></script>
<script defer src="/static/js/mobile-tables.js"></script>
<script defer src="/static/js/validation.js"></script>
<script defer src="/static/js/combobox.js"></script>
<script defer src="/static/js/registry-retry.js"></script>
<script defer src="/static/js/export-summary.js"></script>
<script defer src="/static/js/diff-view.js"></script>
<script defer src="/static/js/undo.js"></script>
<script defer src="/static/js/shortcuts.js"></script>
<script defer src="/static/js/field-help.js"></script>
<script defer src="/static/js/dropzone.js"></script>
<script defer src="/static/js/recent-files.js"></script>
<script defer src="/static/js/onboarding.js"></script>
<script defer src="/static/js/import-profiles.js"></script>
<script defer src="/static/js/bulk-ops.js"></script>
<script defer src="/static/js/error-row-jump.js"></script>
<script defer src="/static/js/i18n.js"></script>
<script defer src="/static/js/audit-log.js"></script>
<script defer src="/static/js/pwa-register.js"></script>
<script>
// Hide empty state once a sample table is rendered
document.addEventListener("DOMContentLoaded", () => {
  const empty = document.getElementById("resultEmptyState");
  const sample = document.getElementById("sampleArea");
  if (!empty || !sample) return;
  const obs = new MutationObserver(() => {
    const hasContent = sample.querySelector("table") || sample.children.length > 0 && !sample.querySelector("p");
    empty.style.display = hasContent ? "none" : "";
  });
  obs.observe(sample, { childList: true, subtree: true });
});
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


class UploadItem:
    def __init__(self, value: str | bytes = "", filename: str | None = None) -> None:
        self.value = value
        self.filename = filename
        raw = value if isinstance(value, bytes) else value.encode("utf-8")
        self.file = io.BytesIO(raw)


class UploadForm(dict[str, UploadItem | list[UploadItem]]):
    def add(self, name: str, item: UploadItem) -> None:
        existing = self.get(name)
        if existing is None:
            self[name] = item
        elif isinstance(existing, list):
            existing.append(item)
        else:
            self[name] = [existing, item]


def get_upload(handler: BaseHTTPRequestHandler) -> UploadForm:
    length = int(handler.headers.get("Content-Length", "0") or 0)
    body = handler.rfile.read(length)
    content_type = handler.headers.get("Content-Type", "")
    form = UploadForm()

    if content_type.startswith("multipart/form-data"):
        message = (
            f"Content-Type: {content_type}\r\n"
            "MIME-Version: 1.0\r\n\r\n"
        ).encode("utf-8") + body
        parsed = BytesParser(policy=email_policy).parsebytes(message)
        for part in parsed.iter_parts():
            name = part.get_param("name", header="content-disposition")
            if not name:
                continue
            filename = part.get_filename()
            payload = part.get_payload(decode=True) or b""
            if filename:
                form.add(name, UploadItem(payload, filename))
            else:
                charset = part.get_content_charset() or "utf-8"
                form.add(name, UploadItem(payload.decode(charset, errors="replace")))
        return form

    if content_type.startswith("application/x-www-form-urlencoded"):
        for key, values in parse_qs(body.decode("utf-8", errors="replace"), keep_blank_values=True).items():
            for value in values:
                form.add(key, UploadItem(value))
    return form


def field_value(form: UploadForm, name: str, default: str = "") -> str:
    item = form[name] if name in form else None
    if item is None:
        return default
    if isinstance(item, list):
        item = item[0]
    if getattr(item, "filename", None):
        return default
    return str(item.value or default)


def upload_bytes(form: UploadForm) -> tuple[str, bytes]:
    if "file" not in form:
        raise ValueError("Hiányzik a feltöltött fájl.")
    item = form["file"]
    if isinstance(item, list):
        item = item[0]
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


def amount_to_fixed(value: str, decimals: int, length: int) -> str:
    raw = clean_cell(value).replace(" ", "").replace("\u00a0", "")
    if not raw:
        raise ValueError("Hiányzik az összeg.")
    raw = raw.replace(",", ".")
    try:
        decimal_amount = Decimal(raw)
    except InvalidOperation as exc:
        raise ValueError(f"Hibás összeg: {value}") from exc
    multiplier = Decimal(10) ** decimals
    number = int((decimal_amount * multiplier).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    if number < 0:
        raise ValueError("Az összeg nem lehet negatív.")
    output = str(number).zfill(length)
    if len(output) > length:
        raise ValueError(f"Az összeg túl hosszú a {length} karakteres mezőhöz: {value}")
    return output


def amount_with_dot(value: str, decimals: int, length: int) -> str:
    raw = clean_cell(value).replace(" ", "").replace("\u00a0", "").replace(",", ".")
    if not raw:
        raise ValueError("Hiányzik az összeg.")
    try:
        decimal_amount = Decimal(raw)
    except InvalidOperation as exc:
        raise ValueError(f"Hibás összeg: {value}") from exc
    quant = Decimal("1").scaleb(-decimals)
    formatted = f"{decimal_amount.quantize(quant, rounding=ROUND_HALF_UP):.{decimals}f}"
    if decimal_amount < 0:
        raise ValueError("Az összeg nem lehet negatív.")
    if len(formatted) > length:
        raise ValueError(f"Az összeg túl hosszú a {length} karakteres mezőhöz: {value}")
    return formatted.rjust(length, "0")


def decimal_amount_string(value: str, decimals: int = 2) -> str:
    raw = clean_cell(value).replace(" ", "").replace("\u00a0", "").replace(",", ".")
    if not raw:
        raise ValueError("Hiányzik az összeg.")
    try:
        decimal_amount = Decimal(raw)
    except InvalidOperation as exc:
        raise ValueError(f"Hibás összeg: {value}") from exc
    if decimal_amount < 0:
        raise ValueError("Az összeg nem lehet negatív.")
    quant = Decimal("1").scaleb(-decimals)
    return f"{decimal_amount.quantize(quant, rounding=ROUND_HALF_UP):.{decimals}f}"


def fixed_local_account(value: str, length: int = 24) -> str:
    digits = digits_only(value)
    if len(digits) not in {16, 24}:
        raise ValueError("A belföldi számlaszám 16 vagy 24 számjegy legyen.")
    return digits.ljust(length, " ")


def checksum_number(value: int, length: int) -> str:
    return str(value)[-length:].zfill(length)


def checksum_text_number(value: str, length: int) -> int:
    normalized = "".join(ch if ch.isdigit() else "0" for ch in value)
    return int(normalized or "0") % (10 ** length)


def optional_date_or_zero(value: str) -> str:
    if not clean_cell(value):
        return "00000000"
    return parse_date(value, "értéknap")


def yn_flag(value: str, default: str = "N") -> str:
    raw = clean_cell(value).upper()
    if raw in {"", "0", "N", "NO", "NEM", "FALSE"}:
        return default if raw == "" else "N"
    if raw in {"1", "Y", "YES", "IGEN", "TRUE"}:
        return "Y"
    raise ValueError(f"Y/N értéket vártam, ezt kaptam: {value}")


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


def build_sepa_payord_line(values: dict[str, str], sequence: int, config: dict[str, Any]) -> str:
    line = [" "] * 1046
    decimals = clean_cell(values.get("decimals") or "2")
    payout_currency = clean_cell(values.get("payout_currency") or "EUR").upper()
    currency = clean_cell(values.get("currency") or "EUR").upper()
    if payout_currency != "EUR" or currency != "EUR":
        raise ValueError("SEPA átutalásnál a kifizetendő és az átutalandó deviza is EUR kell legyen.")
    iban_result = validate_iban_syntax(values.get("beneficiary_account", ""))
    if not iban_result["ok"]:
        raise ValueError("Címzett IBAN: " + "; ".join(iban_result["errors"]))

    identifier = clean_cell(values.get("identifier", ""))
    if not identifier:
        identifier_date = parse_date(config.get("identifier_date") or date.today().strftime("%Y-%m-%d"), "azonosító dátuma")
        identifier = f"{identifier_date}{sequence:06d}"
    identifier = digits_only(identifier)
    if len(identifier) != 14:
        raise ValueError("Az azonosító 14 számjegy legyen, vagy hagyd üresen az automatikus generáláshoz.")

    urgent_execution = clean_cell(values.get("urgent_execution", ""))
    set_field(line, 1, 6, "PAYORD")
    set_field(line, 7, 2, "IN")
    set_field(line, 9, 14, identifier)
    set_field(line, 23, 47, fit_text(values.get("sender_account", ""), 47))
    set_field(line, 70, 1, fit_text(values.get("sender_account_type") or "0", 1))
    set_field(line, 71, 140, fit_text(values.get("sender_name", ""), 140))
    set_field(line, 220, 47, fit_text(iban_result["iban"], 47))
    set_field(line, 267, 1, fit_text(values.get("beneficiary_account_type") or "0", 1))
    set_field(line, 332, 140, fit_text(values.get("beneficiary_name", ""), 140))
    set_field(line, 474, 14, fit_text(values.get("beneficiary_bank_swift", ""), 14))
    set_field(line, 593, 96, fit_text(values.get("note", ""), 96))
    set_field(line, 785, 1, "Y" if urgent_execution else "N")
    set_field(line, 786, 1, fit_text(urgent_execution, 1))
    set_field(line, 789, 3, "EUR")
    set_field(line, 806, 3, "EUR")
    set_field(line, 809, 1, fit_digits(decimals, 1))
    set_field(line, 810, 13, amount_to_field(values.get("amount", ""), decimals))
    set_field(line, 938, 2, "00")
    return "".join(line)


def build_kh_fx_payord_line(values: dict[str, str], sequence: int, config: dict[str, Any]) -> str:
    line = [" "] * 939
    decimals = clean_cell(values.get("decimals") or "2")
    identifier = clean_cell(values.get("identifier", ""))
    if not identifier:
        identifier_date = parse_date(config.get("identifier_date") or date.today().strftime("%Y-%m-%d"), "azonosító dátuma")
        identifier = f"{identifier_date}{sequence:06d}"
    identifier = digits_only(identifier)
    if len(identifier) != 14:
        raise ValueError("Az azonosító 14 számjegy legyen, vagy hagyd üresen az automatikus generáláshoz.")

    set_field(line, 1, 6, "PAYORD")
    set_field(line, 7, 2, "IN")
    set_field(line, 9, 14, identifier)
    set_field(line, 23, 47, fit_text(values.get("sender_account", ""), 47))
    set_field(line, 70, 1, " ")
    set_field(line, 71, 32, fit_text(values.get("sender_name", ""), 32))
    set_field(line, 220, 47, fit_text(values.get("beneficiary_account", ""), 47))
    set_field(line, 267, 1, " ")
    set_field(line, 268, 35, fit_text(values.get("beneficiary_bank_name", ""), 35))
    set_field(line, 303, 29, fit_text(values.get("beneficiary_bank_address", ""), 29))
    set_field(line, 332, 105, fit_text(values.get("beneficiary_name", ""), 105))
    set_field(line, 474, 15, fit_text(values.get("beneficiary_bank_swift", ""), 15))
    set_field(line, 558, 35, fit_text(values.get("bank_message", ""), 35))
    set_field(line, 593, 96, fit_text(values.get("note", ""), 96))
    set_field(line, 789, 3, fit_text(values.get("payout_currency", "").upper(), 3))
    set_field(line, 806, 3, fit_text(values.get("currency", "").upper(), 3))
    set_field(line, 809, 1, fit_digits(decimals, 1))
    set_field(line, 810, 13, amount_to_field(values.get("amount", ""), decimals))
    set_field(line, 835, 8, optional_date(values.get("value_date", ""), "valutanap"))
    set_field(line, 855, 1, fit_text(values.get("cost_bearer") or "1", 1))
    set_field(line, 888, 1, yn_flag(values.get("urgent_execution", ""), "N"))
    set_field(line, 889, 1, yn_flag(values.get("group_transfer", ""), "N"))
    set_field(line, 938, 2, "00")
    return "".join(line)


def build_otp_huf_payord_line(values: dict[str, str], sequence: int, config: dict[str, Any]) -> str:
    line = [" "] * 939
    decimals = clean_cell(values.get("decimals") or "0")
    identifier = clean_cell(values.get("identifier", ""))
    if not identifier:
        identifier_date = parse_date(config.get("identifier_date") or date.today().strftime("%Y-%m-%d"), "azonosító dátuma")
        identifier = f"{identifier_date}{sequence:06d}"
    identifier = digits_only(identifier)
    if len(identifier) != 14:
        raise ValueError("Az azonosító 14 számjegy legyen, vagy hagyd üresen az automatikus generáláshoz.")

    process_mode = clean_cell(values.get("process_mode", "")).upper()
    if process_mode in {"NORMAL", "N", "0"}:
        process_mode = " "
    if process_mode not in {"", " ", "V", "T", "W", "P"}:
        raise ValueError("OTP feldolgozási mód csak üres, V, T, W vagy P lehet.")

    set_field(line, 1, 6, "PAYORD")
    set_field(line, 7, 2, "DO")
    set_field(line, 9, 14, identifier)
    set_field(line, 23, 47, fit_digits(values.get("sender_account", ""), 47))
    set_field(line, 70, 1, "0")
    set_field(line, 71, 32, fit_text(values.get("sender_name", ""), 32))
    set_field(line, 220, 47, fit_digits(values.get("beneficiary_account", ""), 47))
    set_field(line, 267, 1, "0")
    set_field(line, 332, 32, fit_text(values.get("beneficiary_name", ""), 32))
    set_field(line, 472, 2, "  ")
    set_field(line, 593, 96, fit_text(values.get("note", ""), 96))
    set_field(line, 689, 6, fit_text(values.get("document_no", ""), 6))
    set_field(line, 695, 4, "    ")
    set_field(line, 806, 3, fit_text(values.get("currency", "").upper(), 3))
    set_field(line, 809, 1, fit_digits(decimals, 1))
    set_field(line, 810, 13, amount_to_field(values.get("amount", ""), decimals))
    set_field(line, 835, 8, optional_date(values.get("value_date", ""), "értéknap"))
    set_field(line, 843, 1, process_mode or " ")
    set_field(line, 938, 2, "00")
    return "".join(line)


def build_otp_fx_payord_line(values: dict[str, str], sequence: int, config: dict[str, Any]) -> str:
    line = [" "] * 939
    decimals = clean_cell(values.get("decimals") or "2")
    identifier = clean_cell(values.get("identifier", ""))
    if not identifier:
        identifier_date = parse_date(config.get("identifier_date") or date.today().strftime("%Y-%m-%d"), "azonosító dátuma")
        identifier = f"{identifier_date}{sequence:06d}"
    identifier = digits_only(identifier)
    if len(identifier) != 14:
        raise ValueError("Az azonosító 14 számjegy legyen, vagy hagyd üresen az automatikus generáláshoz.")

    urgent = clean_cell(values.get("urgent_execution", ""))
    if urgent and urgent not in {"0", "1", "2"}:
        raise ValueError("OTP sürgős teljesítés mező csak üres, 0, 1 vagy 2 lehet.")

    set_field(line, 1, 6, "PAYORD")
    set_field(line, 7, 2, "IN")
    set_field(line, 9, 14, identifier)
    set_field(line, 23, 47, fit_text(values.get("sender_account", ""), 47))
    set_field(line, 70, 1, fit_text(values.get("sender_account_type") or "0", 1))
    set_field(line, 71, 140, fit_text(values.get("sender_name", ""), 140))
    set_field(line, 220, 47, fit_text(values.get("beneficiary_account", ""), 47))
    set_field(line, 267, 1, fit_text(values.get("beneficiary_account_type") or "0", 1))
    set_field(line, 268, 35, fit_text(values.get("beneficiary_bank_name", ""), 35))
    set_field(line, 303, 29, fit_text(values.get("beneficiary_bank_address", ""), 29))
    set_field(line, 332, 140, fit_text(values.get("beneficiary_name", ""), 140))
    set_field(line, 474, 14, fit_text(values.get("beneficiary_bank_swift", ""), 14))
    set_field(line, 489, 35, fit_text(values.get("correspondent_bank_1", ""), 35))
    set_field(line, 558, 35, fit_text(values.get("bank_message", ""), 35))
    set_field(line, 593, 96, fit_text(values.get("note", ""), 96))
    set_field(line, 753, 24, fit_text(values.get("permit_no", ""), 24))
    set_field(line, 777, 8, optional_date(values.get("permit_date", ""), "deviza engedély dátuma"))
    set_field(line, 786, 1, urgent or " ")
    set_field(line, 789, 3, fit_text(values.get("payout_currency", "").upper(), 3))
    set_field(line, 806, 3, fit_text(values.get("currency", "").upper(), 3))
    set_field(line, 809, 1, fit_digits(decimals, 1))
    set_field(line, 810, 13, amount_to_field(values.get("amount", ""), decimals))
    set_field(line, 855, 1, fit_text(values.get("cost_bearer") or "1", 1))
    set_field(line, 938, 2, "00")
    return "".join(line)


def build_raiffeisen_huf_line(values: dict[str, str], sequence: int, config: dict[str, Any]) -> str:
    line = [" "] * 249
    set_field(line, 1, 15, amount_with_dot(values.get("amount", ""), 2, 15))
    set_field(line, 16, 8, optional_date(values.get("value_date", ""), "értéknap"))
    set_field(line, 24, 13, " " * 13)
    set_field(line, 37, 24, fixed_local_account(values.get("sender_account", ""), 24))
    set_field(line, 61, 24, fixed_local_account(values.get("beneficiary_account", ""), 24))
    set_field(line, 85, 32, fit_text(values.get("beneficiary_name", ""), 32))
    set_field(line, 117, 2, fit_text(values.get("beneficiary_country", ""), 2))
    set_field(line, 119, 3, fit_text(values.get("legal_code", ""), 3))
    set_field(line, 137, 70, fit_text(values.get("note", ""), 70))
    set_field(line, 213, 10, fit_text(values.get("external_ref", ""), 10))
    set_field(line, 223, 6, fit_text(values.get("document_no", ""), 6))
    set_field(line, 229, 20, fit_text(values.get("internal_note", ""), 20))
    return "".join(line)


def build_raiffeisen_fx_line(values: dict[str, str], sequence: int, config: dict[str, Any]) -> str:
    line = [" "] * 493
    note = fit_text(values.get("note", ""), 120)
    beneficiary_name = fit_text(values.get("beneficiary_name", ""), 105)
    set_field(line, 1, 3, fit_text(values.get("currency", "").upper(), 3))
    set_field(line, 4, 3, fit_text(values.get("sender_currency", "").upper(), 3))
    set_field(line, 7, 15, amount_with_dot(values.get("amount", ""), 3, 15))
    set_field(line, 22, 8, optional_date(values.get("value_date", ""), "értéknap"))
    set_field(line, 30, 13, fit_digits(values.get("sender_account", ""), 13))
    set_field(line, 43, 30, beneficiary_name[0:30])
    set_field(line, 73, 30, beneficiary_name[30:60])
    set_field(line, 103, 30, beneficiary_name[60:90])
    set_field(line, 133, 15, beneficiary_name[90:105])
    set_field(line, 148, 30, fit_text(values.get("beneficiary_bank_name", ""), 30))
    set_field(line, 178, 30, " " * 30)
    set_field(line, 208, 26, fit_text(values.get("beneficiary_bank_country", ""), 26))
    set_field(line, 234, 30, fit_text(values.get("beneficiary_bank_address_1", ""), 30))
    set_field(line, 264, 20, fit_text(values.get("beneficiary_bank_address_2", ""), 20))
    set_field(line, 284, 35, fit_text(values.get("beneficiary_account", ""), 35))
    set_field(line, 319, 4, fit_text(values.get("legal_code", ""), 4))
    set_field(line, 323, 30, note[0:30])
    set_field(line, 353, 30, note[30:60])
    set_field(line, 383, 30, note[60:90])
    set_field(line, 413, 30, note[90:120])
    set_field(line, 443, 1, fit_text(values.get("commission_bearer") or "0", 1))
    set_field(line, 444, 1, fit_text(values.get("other_fee_bearer") or "0", 1))
    set_field(line, 445, 10, fit_text(values.get("permit_no", ""), 10))
    set_field(line, 455, 6, fit_text(values.get("document_no", ""), 6))
    set_field(line, 461, 11, fit_text(values.get("beneficiary_bank_swift", ""), 11))
    set_field(line, 472, 1, fit_text(values.get("amount_currency_mode") or " ", 1))
    set_field(line, 473, 1, fit_text(values.get("payment_method") or " ", 1))
    set_field(line, 474, 1, " ")
    set_field(line, 475, 1, fit_text(values.get("priority") or " ", 1))
    set_field(line, 476, 1, fit_text(values.get("item_type") or " ", 1))
    set_field(line, 477, 1, fit_text(values.get("iban_flag") or " ", 1))
    set_field(line, 478, 2, fit_text(values.get("beneficiary_country", ""), 2))
    set_field(line, 480, 3, "   ")
    set_field(line, 483, 10, fit_text(values.get("external_ref", ""), 10))
    return "".join(line)


def xml_text(parent: ET.Element, tag: str, text: str | None = None) -> ET.Element:
    child = ET.SubElement(parent, tag)
    if text is not None:
        child.text = text
    return child


def pain_account(parent: ET.Element, account: str) -> None:
    acct_id = xml_text(parent, "Id")
    cleaned = clean_cell(account).replace(" ", "")
    if re.match(r"^[A-Z]{2}[0-9A-Z]+$", cleaned.upper()):
        xml_text(acct_id, "IBAN", cleaned.upper())
    else:
        othr = xml_text(acct_id, "Othr")
        xml_text(othr, "Id", digits_only(cleaned) or cleaned)


def build_pain_xml(records: list[dict[str, str]], config: dict[str, Any], service_level: str | None = None, version: str = "pain.001.001.03") -> str:
    if not records:
        raise ValueError("Nem találtam konvertálható adatsort.")
    ns = f"urn:iso:std:iso:20022:tech:xsd:{version}"
    ET.register_namespace("", ns)
    doc = ET.Element(f"{{{ns}}}Document")
    cstmr = xml_text(doc, "CstmrCdtTrfInitn")
    now = datetime.now().replace(microsecond=0).isoformat()
    msg_id = clean_cell(records[0].get("message_id")) or f"MSG{datetime.now().strftime('%Y%m%d%H%M%S')}"
    total = sum(Decimal(decimal_amount_string(r.get("amount", ""), 2)) for r in records)

    grp = xml_text(cstmr, "GrpHdr")
    xml_text(grp, "MsgId", msg_id)
    xml_text(grp, "CreDtTm", now)
    xml_text(grp, "NbOfTxs", str(len(records)))
    xml_text(grp, "CtrlSum", f"{total:.2f}")
    initg = xml_text(grp, "InitgPty")
    xml_text(initg, "Nm", clean_cell(records[0].get("debtor_name"))[:70])

    pmt = xml_text(cstmr, "PmtInf")
    xml_text(pmt, "PmtInfId", clean_cell(records[0].get("payment_id")) or f"PMT{datetime.now().strftime('%Y%m%d%H%M%S')}")
    xml_text(pmt, "PmtMtd", "TRF")
    if service_level:
        pti = xml_text(pmt, "PmtTpInf")
        svc = xml_text(pti, "SvcLvl")
        xml_text(svc, "Cd", service_level)
    req_date = parse_date(records[0].get("value_date", ""), "értéknap")
    xml_text(pmt, "ReqdExctnDt", f"{req_date[:4]}-{req_date[4:6]}-{req_date[6:8]}")
    dbtr = xml_text(pmt, "Dbtr")
    xml_text(dbtr, "Nm", clean_cell(records[0].get("debtor_name"))[:70])
    dbtr_acct = xml_text(pmt, "DbtrAcct")
    pain_account(dbtr_acct, records[0].get("sender_account", ""))

    for seq, values in enumerate(records, start=1):
        cdt = xml_text(pmt, "CdtTrfTxInf")
        pmtid = xml_text(cdt, "PmtId")
        xml_text(pmtid, "EndToEndId", clean_cell(values.get("end_to_end_id")) or f"NOTPROVIDED{seq:06d}")
        amt = xml_text(cdt, "Amt")
        instd = xml_text(amt, "InstdAmt", decimal_amount_string(values.get("amount", ""), 2))
        instd.set("Ccy", clean_cell(values.get("currency") or "EUR").upper())
        cb = clean_cell(values.get("cost_bearer") or ("SLEV" if service_level == "SEPA" else "SHAR")).upper()
        xml_text(cdt, "ChrgBr", cb)
        if clean_cell(values.get("beneficiary_bank_swift", "")):
            agent = xml_text(cdt, "CdtrAgt")
            fin = xml_text(agent, "FinInstnId")
            xml_text(fin, "BIC", clean_cell(values.get("beneficiary_bank_swift", "")).upper())
        cdtr = xml_text(cdt, "Cdtr")
        xml_text(cdtr, "Nm", clean_cell(values.get("beneficiary_name", ""))[:70])
        if clean_cell(values.get("beneficiary_country", "")):
            adr = xml_text(cdtr, "PstlAdr")
            xml_text(adr, "Ctry", clean_cell(values.get("beneficiary_country", "")).upper()[:2])
        cdtr_acct = xml_text(cdt, "CdtrAcct")
        pain_account(cdtr_acct, values.get("beneficiary_account", ""))
        if clean_cell(values.get("note", "")):
            rmt = xml_text(cdt, "RmtInf")
            xml_text(rmt, "Ustrd", clean_cell(values.get("note", ""))[:140])

    raw = ET.tostring(doc, encoding="utf-8")
    pretty = minidom.parseString(raw).toprettyxml(indent="  ", encoding="utf-8")
    return pretty.decode("utf-8")


def build_unicredit_pay_text(records: list[dict[str, str]], config: dict[str, Any]) -> str:
    if not records:
        raise ValueError("Nem találtam konvertálható adatsort.")
    first = records[0]
    creation_date = parse_date(config.get("identifier_date") or date.today().strftime("%Y-%m-%d"), "létrehozás dátuma")
    sender_account = fixed_local_account(first.get("sender_account", ""))
    sender_currency = fit_text(first.get("sender_currency") or "HUF", 3)

    header = [" "] * 256
    set_field(header, 1, 2, "23")
    set_field(header, 3, 24, sender_account)
    set_field(header, 27, 3, sender_currency)
    set_field(header, 30, 1, "1")
    set_field(header, 31, 1, "0")
    set_field(header, 32, 8, creation_date)

    lines = ["".join(header)]
    record_sum = 0
    account_sum = 0
    amount_sum = 0
    for sequence, values in enumerate(records, start=1):
        if fixed_local_account(values.get("sender_account", "")) != sender_account:
            raise ValueError("UniCredit PAY fájlban egy csomagon belül csak egy terhelendő számlaszám szerepelhet.")
        if fit_text(values.get("sender_currency") or "HUF", 3) != sender_currency:
            raise ValueError("UniCredit PAY fájlban egy csomagon belül csak egy terhelendő számladeviza szerepelhet.")
        currency = clean_cell(values.get("currency") or "HUF").upper()
        decimals = 0 if currency in {"HUF", "JPY"} else 2
        account = fixed_local_account(values.get("beneficiary_account", ""))
        amount = amount_to_fixed(values.get("amount", ""), 0, 13) + "00" if decimals == 0 else amount_to_fixed(values.get("amount", ""), 2, 15)

        detail = [" "] * 256
        set_field(detail, 1, 2, "43")
        set_field(detail, 3, 6, checksum_number(sequence, 6))
        set_field(detail, 9, 24, account)
        set_field(detail, 33, 32, fit_text(values.get("beneficiary_name", ""), 32))
        set_field(detail, 65, 6, fit_text(values.get("document_no", ""), 6))
        note = fit_text(values.get("note", ""), 96)
        set_field(detail, 71, 32, note[0:32])
        set_field(detail, 103, 32, note[32:64])
        set_field(detail, 135, 32, note[64:96])
        set_field(detail, 167, 15, amount)
        set_field(detail, 182, 3, fit_text(currency, 3))
        set_field(detail, 185, 8, optional_date_or_zero(values.get("value_date", "")))
        set_field(detail, 193, 3, fit_text(values.get("legal_code", ""), 3))
        set_field(detail, 196, 2, fit_text(values.get("beneficiary_country", ""), 2))
        set_field(detail, 205, 3, "000")
        set_field(detail, 208, 3, "000")
        detail_text = "".join(detail)
        lines.append(detail_text)
        record_sum += sequence
        account_sum += checksum_text_number(account, 24)
        amount_sum += int(amount)

    trailer = [" "] * 256
    set_field(trailer, 1, 2, "63")
    set_field(trailer, 3, 6, checksum_number(record_sum, 6))
    set_field(trailer, 9, 24, checksum_number(account_sum, 24))
    set_field(trailer, 167, 15, checksum_number(amount_sum, 15))
    lines.append("".join(trailer))
    return "\r\n".join(lines) + "\r\n"


def build_unicredit_ccy_text(records: list[dict[str, str]], config: dict[str, Any]) -> str:
    if not records:
        raise ValueError("Nem találtam konvertálható adatsort.")
    first = records[0]
    creation_date = parse_date(config.get("identifier_date") or date.today().strftime("%Y-%m-%d"), "létrehozás dátuma")
    sender_account = fixed_local_account(first.get("sender_account", ""))
    sender_currency = fit_text(first.get("sender_currency") or "HUF", 3)

    header = [" "] * 800
    set_field(header, 1, 2, "34")
    set_field(header, 3, 24, sender_account)
    set_field(header, 27, 3, sender_currency)
    set_field(header, 30, 1, "1")
    set_field(header, 31, 1, "0")
    set_field(header, 32, 8, creation_date)

    lines = ["".join(header)]
    record_sum = 0
    amount_sum = 0
    for sequence, values in enumerate(records, start=1):
        if fixed_local_account(values.get("sender_account", "")) != sender_account:
            raise ValueError("UniCredit CCY fájlban egy csomagon belül csak egy terhelendő számlaszám szerepelhet.")
        if fit_text(values.get("sender_currency") or "HUF", 3) != sender_currency:
            raise ValueError("UniCredit CCY fájlban egy csomagon belül csak egy terhelendő számladeviza szerepelhet.")
        currency = clean_cell(values.get("currency") or "EUR").upper()
        decimals = 0 if currency in {"HUF", "JPY"} else 2
        amount = amount_to_fixed(values.get("amount", ""), 0, 13) + "00" if decimals == 0 else amount_to_fixed(values.get("amount", ""), 2, 15)

        note = fit_text(values.get("note", ""), 140)
        detail = [" "] * 800
        set_field(detail, 1, 2, "54")
        set_field(detail, 3, 6, checksum_number(sequence, 6))
        set_field(detail, 9, 11, fit_text(values.get("beneficiary_bank_swift", ""), 11))
        set_field(detail, 20, 33, fit_text(values.get("beneficiary_bank_id", ""), 33))
        set_field(detail, 53, 35, fit_text(values.get("beneficiary_bank_name", ""), 35))
        set_field(detail, 88, 35, fit_text(values.get("beneficiary_bank_address_1", ""), 35))
        set_field(detail, 123, 35, fit_text(values.get("beneficiary_bank_address_2", ""), 35))
        set_field(detail, 158, 35, fit_text(values.get("beneficiary_bank_address_3", ""), 35))
        set_field(detail, 193, 11, fit_text(values.get("correspondent_bank_swift", ""), 11))
        set_field(detail, 204, 33, fit_text(values.get("correspondent_bank_id", ""), 33))
        set_field(detail, 237, 35, fit_text(values.get("correspondent_bank_name", ""), 35))
        set_field(detail, 272, 35, fit_text(values.get("correspondent_bank_address_1", ""), 35))
        set_field(detail, 307, 35, fit_text(values.get("correspondent_bank_address_2", ""), 35))
        set_field(detail, 342, 35, fit_text(values.get("correspondent_bank_address_3", ""), 35))
        set_field(detail, 377, 34, fit_text(values.get("beneficiary_account", ""), 34))
        set_field(detail, 411, 35, fit_text(values.get("beneficiary_name", ""), 35))
        set_field(detail, 446, 35, fit_text(values.get("beneficiary_address_1", ""), 35))
        set_field(detail, 481, 35, fit_text(values.get("beneficiary_address_2", ""), 35))
        set_field(detail, 516, 35, fit_text(values.get("beneficiary_address_3", ""), 35))
        set_field(detail, 551, 35, note[0:35])
        set_field(detail, 586, 35, note[35:70])
        set_field(detail, 621, 35, note[70:105])
        set_field(detail, 656, 35, note[105:140])
        set_field(detail, 691, 3, fit_text(values.get("cost_bearer") or "SHA", 3))
        set_field(detail, 694, 3, fit_text(values.get("payout_currency") or currency, 3))
        set_field(detail, 697, 15, amount)
        set_field(detail, 712, 3, fit_text(currency, 3))
        set_field(detail, 715, 8, optional_date_or_zero(values.get("value_date", "")))
        set_field(detail, 723, 1, yn_flag(values.get("urgent_execution", ""), "N"))
        set_field(detail, 724, 4, "0000")
        set_field(detail, 728, 1, yn_flag(values.get("hold_flag", ""), "N"))
        set_field(detail, 729, 1, yn_flag(values.get("chqb_flag", ""), "N"))
        set_field(detail, 730, 1, yn_flag(values.get("deal_ticket_flag", ""), "N"))
        set_field(detail, 731, 8, optional_date_or_zero(values.get("deal_ticket_date", "")))
        set_field(detail, 739, 6, fit_digits(values.get("deal_ticket_sequence", ""), 6, pad_left=True))
        set_field(detail, 745, 2, fit_text(values.get("beneficiary_country", ""), 2))
        set_field(detail, 747, 3, fit_text(values.get("legal_code", ""), 3))
        set_field(detail, 750, 3, "000")
        set_field(detail, 753, 3, "000")
        lines.append("".join(detail))
        record_sum += sequence
        amount_sum += int(amount)

    trailer = [" "] * 800
    set_field(trailer, 1, 2, "74")
    set_field(trailer, 3, 6, checksum_number(record_sum, 6))
    set_field(trailer, 697, 15, checksum_number(amount_sum, 15))
    lines.append("".join(trailer))
    return "\r\n".join(lines) + "\r\n"


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
    format_def = get_format(format_id)
    if format_def.get("unsupported"):
        raise ValueError(format_def["unsupported"])
    fields = fields_for_format(format_id)
    sheets, selected, headers, rows = workbook_rows(data, filename)
    if not headers:
        raise ValueError("Nem található fejlécsor.")
    output: list[str] = []
    structured_records: list[dict[str, str]] = []
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
            builder = get_format(format_id).get("builder")
            if format_id in {"unicredit_huf_pay", "unicredit_fx_ccy"} or builder in {"pain_huf", "pain_fx", "pain_sepa"}:
                structured_records.append(values)
                continue
            if format_id == "erste_fx_payord":
                line = build_fx_payord_line(values, sequence, config)
            elif format_id == "erste_sepa_payord":
                line = build_sepa_payord_line(values, sequence, config)
            elif format_id == "kh_fx_payord":
                line = build_kh_fx_payord_line(values, sequence, config)
            elif format_id == "otp_huf":
                line = build_otp_huf_payord_line(values, sequence, config)
            elif format_id == "otp_fx":
                line = build_otp_fx_payord_line(values, sequence, config)
            elif format_id == "raiffeisen_huf":
                line = build_raiffeisen_huf_line(values, sequence, config)
            elif format_id in {"raiffeisen_fx", "raiffeisen_sepa"}:
                line = build_raiffeisen_fx_line(values, sequence, config)
            else:
                line = build_payord_line(values, sequence, config)
            output.append(line)
        except ValueError as exc:
            errors.append(f"{offset}. sor: {exc}")
            if len(errors) >= 12:
                break
    if errors:
        raise ValueError("\n".join(errors))
    if format_id == "unicredit_huf_pay":
        return build_unicredit_pay_text(structured_records, config), len(structured_records)
    if format_id == "unicredit_fx_ccy":
        return build_unicredit_ccy_text(structured_records, config), len(structured_records)
    builder = get_format(format_id).get("builder")
    if builder == "pain_sepa":
        return build_pain_xml(structured_records, config, service_level="SEPA"), len(structured_records)
    if builder in {"pain_huf", "pain_fx"}:
        return build_pain_xml(structured_records, config), len(structured_records)
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


HU_WEIGHTS = [9, 7, 3, 1]
IBAN_LENGTHS = {
    "AD": 24, "AE": 23, "AL": 28, "AT": 20, "AZ": 28, "BA": 20, "BE": 16, "BG": 22,
    "BH": 22, "BR": 29, "BY": 28, "CH": 21, "CR": 22, "CY": 28, "CZ": 24, "DE": 22,
    "DK": 18, "DO": 28, "EE": 20, "EG": 29, "ES": 24, "FI": 18, "FO": 18, "FR": 27,
    "GB": 22, "GE": 22, "GI": 23, "GL": 18, "GR": 27, "GT": 28, "HR": 21, "HU": 28,
    "IE": 22, "IL": 23, "IQ": 23, "IS": 26, "IT": 27, "JO": 30, "KW": 30, "KZ": 20,
    "LB": 28, "LC": 32, "LI": 21, "LT": 20, "LU": 20, "LV": 21, "MC": 27, "MD": 24,
    "ME": 22, "MK": 19, "MR": 27, "MT": 31, "MU": 30, "NL": 18, "NO": 15, "PK": 24,
    "PL": 28, "PS": 29, "PT": 25, "QA": 29, "RO": 24, "RS": 22, "SA": 24, "SC": 31,
    "SE": 24, "SI": 19, "SK": 24, "SM": 27, "ST": 25, "SV": 28, "TL": 23, "TN": 24,
    "TR": 26, "UA": 29, "VA": 22, "VG": 24, "XK": 20,
}


def default_bank_registry() -> dict[str, Any]:
    return {
        "source_url": "https://www.mnb.hu/letoltes/sht.xlsx",
        "name_column": "C",
        "prefix_column": "A",
        "updated_at": "",
        "rows": [],
        "detected_name_column": "",
        "detected_prefix_column": "",
        "startup_ok": None,
        "startup_message": "",
    }


def load_json_file(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default


def write_json_file(path: Path, data: Any) -> Any:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return data


def load_accounts_file() -> list[dict[str, Any]]:
    data = load_json_file(ACCOUNTS_FILE, [])
    if isinstance(data, dict):
        rows: list[dict[str, Any]] = []
        for company_id, accounts in (data.get("by_company") or {}).items():
            if isinstance(accounts, list):
                for row in accounts:
                    if isinstance(row, dict):
                        rows.append({**row, "company_id": clean_cell(row.get("company_id") or company_id or "default")})
        data = rows
    if not isinstance(data, list):
        return []
    for row in data:
        if isinstance(row, dict):
            row["currency"] = clean_cell(row.get("currency") or "HUF").upper()
            row["company_id"] = clean_cell(row.get("company_id") or "default")
    return data


def save_accounts_file(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return write_json_file(ACCOUNTS_FILE, rows)


def default_company() -> dict[str, str]:
    return {"id": "default", "name": "Alap cég"}


def load_companies_file() -> list[dict[str, Any]]:
    data = load_json_file(COMPANIES_FILE, [])
    if not isinstance(data, list) or not data:
        return [default_company()]
    normalized: list[dict[str, Any]] = []
    seen: set[str] = set()
    for row in data:
        if not isinstance(row, dict):
            continue
        company_id = clean_cell(row.get("id")) or uuid.uuid4().hex
        name = clean_cell(row.get("name")) or "Névtelen cég"
        if company_id in seen:
            continue
        normalized.append({"id": company_id, "name": name})
        seen.add(company_id)
    if "default" not in seen:
        normalized.insert(0, default_company())
    return normalized


def save_companies_file(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return write_json_file(COMPANIES_FILE, rows)


def ensure_company_id(company_id: str | None = None) -> str:
    requested = clean_cell(company_id or "default")
    companies = load_companies_file()
    if any(company.get("id") == requested for company in companies):
        return requested
    return companies[0]["id"] if companies else "default"


def accounts_for_company(company_id: str | None) -> list[dict[str, Any]]:
    cid = ensure_company_id(company_id)
    return [row for row in load_accounts_file() if clean_cell(row.get("company_id") or "default") == cid]


def load_partners_file() -> list[dict[str, Any]]:
    data = load_json_file(PARTNERS_FILE, [])
    if isinstance(data, dict):
        rows: list[dict[str, Any]] = []
        for company_id, partners in (data.get("by_company") or {}).items():
            if isinstance(partners, list):
                for row in partners:
                    if isinstance(row, dict):
                        rows.append({**row, "company_id": clean_cell(row.get("company_id") or company_id or "default")})
        data = rows
    if not isinstance(data, list):
        return []
    rows = [row for row in data if isinstance(row, dict)]
    for row in rows:
        row["company_id"] = clean_cell(row.get("company_id") or "default")
    return rows


def save_partners_file(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return write_json_file(PARTNERS_FILE, rows)


def partners_for_company(company_id: str | None) -> list[dict[str, Any]]:
    cid = ensure_company_id(company_id)
    return [row for row in load_partners_file() if clean_cell(row.get("company_id") or "default") == cid]


def load_bank_registry() -> dict[str, Any]:
    data = load_json_file(BANK_REGISTRY_FILE, default_bank_registry())
    base = default_bank_registry()
    if isinstance(data, dict):
        base.update(data)
    if not isinstance(base.get("rows"), list):
        base["rows"] = []
    return base


def save_bank_registry(data: dict[str, Any]) -> dict[str, Any]:
    base = load_bank_registry()
    for key in ("source_url", "name_column", "prefix_column"):
        if key in data:
            base[key] = clean_cell(data.get(key))
    return write_json_file(BANK_REGISTRY_FILE, base)


def strip_accents(value: str) -> str:
    text = unicodedata.normalize("NFD", str(value or ""))
    return "".join(ch for ch in text if unicodedata.category(ch) != "Mn")


def loose_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", strip_accents(value).lower())


def hu_check_digit(payload: str) -> int:
    total = sum(int(ch) * HU_WEIGHTS[index % len(HU_WEIGHTS)] for index, ch in enumerate(payload))
    return (10 - (total % 10)) % 10


def format_hu_account_digits(digits: str) -> str:
    padded = digits if len(digits) == 24 else digits + "0" * 8
    return "-".join(padded[index:index + 8] for index in range(0, 24, 8))


def registry_lookup(prefix: str, registry: dict[str, Any] | None = None) -> dict[str, str] | None:
    registry = registry or load_bank_registry()
    for row in registry.get("rows", []):
        if clean_cell(row.get("prefix")) == prefix:
            return row
    return None


def validate_hu_account(value: str, registry: dict[str, Any] | None = None) -> dict[str, Any]:
    digits = digits_only(value)
    errors: list[str] = []
    if len(digits) not in (16, 24):
        errors.append("A magyar bankszámlaszám 16 vagy 24 számjegy lehet.")
        return {"ok": False, "errors": errors, "digits": digits, "formatted": "", "bank": None}
    if hu_check_digit(digits[:7]) != int(digits[7]):
        errors.append("Az első 8 számjegy ellenőrzőszáma hibás.")
    second_payload = digits[8:-1]
    if hu_check_digit(second_payload) != int(digits[-1]):
        errors.append("A számlaszám 2-3. nyolcas csoportjának ellenőrzőszáma hibás.")
    bank = registry_lookup(digits[:8], registry)
    if not bank:
        errors.append("Az első 8 számjegy nincs a betöltött hitelesítő táblában.")
    return {
        "ok": not errors,
        "errors": errors,
        "digits": digits,
        "normalized": digits if len(digits) == 24 else digits + "0" * 8,
        "formatted": format_hu_account_digits(digits),
        "prefix": digits[:8],
        "bank": bank,
    }


def normalize_iban(value: str) -> str:
    return re.sub(r"[\s-]+", "", clean_cell(value)).upper()


def iban_mod97_ok(iban: str) -> bool:
    remainder = 0
    for ch in iban[4:] + iban[:4]:
        if ch.isdigit():
            expanded = ch
        elif "A" <= ch <= "Z":
            expanded = str(ord(ch) - 55)
        else:
            return False
        for digit in expanded:
            remainder = (remainder * 10 + int(digit)) % 97
    return remainder == 1


def validate_iban(value: str, registry: dict[str, Any] | None = None) -> dict[str, Any]:
    iban = normalize_iban(value)
    errors: list[str] = []
    if not re.fullmatch(r"[A-Z]{2}[0-9]{2}[A-Z0-9]{1,30}", iban):
        errors.append("Az IBAN formátuma hibás.")
    expected_length = IBAN_LENGTHS.get(iban[:2])
    if expected_length and len(iban) != expected_length:
        errors.append(f"{iban[:2]} IBAN hossza {expected_length} karakter kell legyen.")
    elif len(iban) > 34:
        errors.append("Az IBAN legfeljebb 34 karakter lehet.")
    if not errors and not iban_mod97_ok(iban):
        errors.append("Az IBAN MOD 97-10 ellenőrzése hibás.")
    domestic = None
    if not errors and iban.startswith("HU"):
        domestic = validate_hu_account(iban[4:], registry)
        if not domestic["ok"]:
            errors.extend(f"HU IBAN belföldi rész: {msg}" for msg in domestic["errors"])
    elif not errors:
        errors.append("Külföldi IBAN-t egyelőre nem kezelünk ebben a listában.")
    return {"ok": not errors, "errors": errors, "iban": iban, "domestic": domestic}


def validate_iban_syntax(value: str) -> dict[str, Any]:
    iban = normalize_iban(value)
    errors: list[str] = []
    if not re.fullmatch(r"[A-Z]{2}[0-9]{2}[A-Z0-9]{1,30}", iban):
        errors.append("Az IBAN formátuma hibás.")
    expected_length = IBAN_LENGTHS.get(iban[:2])
    if expected_length and len(iban) != expected_length:
        errors.append(f"{iban[:2]} IBAN hossza {expected_length} karakter kell legyen.")
    elif len(iban) > 34:
        errors.append("Az IBAN legfeljebb 34 karakter lehet.")
    if not errors and not iban_mod97_ok(iban):
        errors.append("Az IBAN MOD 97-10 ellenőrzése hibás.")
    return {"ok": not errors, "errors": errors, "iban": iban}


def validate_account_input(value: str, registry: dict[str, Any] | None = None) -> dict[str, Any]:
    raw = clean_cell(value)
    if re.match(r"^[A-Za-z]{2}", raw.replace(" ", "").replace("-", "")):
        iban_result = validate_iban(raw, registry)
        domestic = iban_result.get("domestic") or {}
        return {
            "ok": iban_result["ok"],
            "errors": iban_result["errors"],
            "input_type": "IBAN",
            "iban": iban_result["iban"],
            "digits": domestic.get("digits", ""),
            "normalized": domestic.get("normalized", ""),
            "formatted": domestic.get("formatted", ""),
            "prefix": domestic.get("prefix", ""),
            "bank": domestic.get("bank"),
        }
    domestic_result = validate_hu_account(raw, registry)
    domestic_result["input_type"] = "HU"
    return domestic_result


def selector_to_index(selector: str, headers: list[str]) -> int | None:
    raw = clean_cell(selector)
    if not raw:
        return None
    if re.fullmatch(r"[A-Za-z]+", raw):
        try:
            return openpyxl.utils.column_index_from_string(raw.upper()) - 1
        except ValueError:
            pass
    if raw.isdigit() and int(raw) > 0:
        return int(raw) - 1
    target = loose_key(raw)
    for index, header in enumerate(headers):
        if loose_key(header) == target:
            return index
    return None


def auto_detect_registry_columns(rows: list[list[Any]]) -> tuple[int, int]:
    scan_rows = rows[:120]
    width = max((len(row) for row in scan_rows), default=0)
    prefix_scores: list[tuple[int, int]] = []
    name_scores: list[tuple[int, int]] = []
    for col in range(width):
        prefix_count = 0
        name_count = 0
        for row in scan_rows[1:]:
            value = clean_cell(row[col] if col < len(row) else "")
            digits = digits_only(value)
            if len(digits) >= 8:
                prefix_count += 1
            if value and len(digits) < max(3, len(value) // 2) and len(value) > 3:
                name_count += 1
        prefix_scores.append((prefix_count, col))
        name_scores.append((name_count, col))
    prefix_col = max(prefix_scores, default=(0, 0))[1]
    name_candidates = [pair for pair in name_scores if pair[1] != prefix_col]
    name_col = max(name_candidates, default=(0, 1 if width > 1 else 0))[1]
    return name_col, prefix_col


def parse_bank_registry(data: bytes, config: dict[str, Any]) -> dict[str, Any]:
    workbook = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    ws = workbook[workbook.sheetnames[0]]
    rows = [list(row) for row in ws.iter_rows(values_only=True)]
    if not rows:
        raise ValueError("A hitelesítő tábla üres.")
    headers = [clean_cell(v) for v in rows[0]]
    detected_name, detected_prefix = auto_detect_registry_columns(rows)
    header_keys = [loose_key(header) for header in headers]
    for index, key in enumerate(header_keys):
        if key in {"branchofficecode", "branchcode"}:
            detected_prefix = index
        if key in {"nameofthebranchoffice", "branchname", "bankname", "name"}:
            detected_name = index
    name_col = selector_to_index(config.get("name_column", ""), headers)
    prefix_col = selector_to_index(config.get("prefix_column", ""), headers)
    if name_col is None:
        name_col = detected_name
    if prefix_col is None:
        prefix_col = detected_prefix
    parsed: dict[str, dict[str, str]] = {}
    for row in rows[1:]:
        name = clean_cell(row[name_col] if name_col < len(row) else "")
        prefix = digits_only(row[prefix_col] if prefix_col < len(row) else "")[:8]
        if len(prefix) == 8 and name:
            parsed[prefix] = {"prefix": prefix, "bank_name": name, "bank_country": "HU"}
    if not parsed:
        raise ValueError("Nem találtam 8 számjegyű prefixet és banknevet a megadott oszlopokkal.")
    return {
        **config,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "detected_name_column": openpyxl.utils.get_column_letter(name_col + 1),
        "detected_prefix_column": openpyxl.utils.get_column_letter(prefix_col + 1),
        "rows": sorted(parsed.values(), key=lambda item: (item["bank_name"], item["prefix"])),
    }


def refresh_bank_registry() -> dict[str, Any]:
    config = load_bank_registry()
    url = clean_cell(config.get("source_url")) or default_bank_registry()["source_url"]
    request = urllib.request.Request(url, headers={"User-Agent": "Banki import konvertalo"})
    with urllib.request.urlopen(request, timeout=30) as response:
        data = response.read()
    parsed = parse_bank_registry(data, config)
    return write_json_file(BANK_REGISTRY_FILE, parsed)


def bank_registry_response(registry: dict[str, Any], *, status: str | None = None, error: str = "") -> dict[str, Any]:
    rows = registry.get("rows", []) or []
    payload = {
        **registry,
        "row_count": len(rows),
        "rows_loaded": bool(rows),
        "sample": rows[:8],
    }
    if status:
        payload["status"] = status
    if error:
        payload["error"] = error
    return payload


def refresh_bank_registry_at_startup() -> dict[str, Any]:
    previous = load_bank_registry()
    try:
        registry = refresh_bank_registry()
        count = len(registry.get("rows", []))
        registry["startup_ok"] = True
        registry["startup_message"] = f"Hitelesítő tábla indításkor frissítve: {count} prefix, időpont: {registry.get('updated_at', '')}."
        return write_json_file(BANK_REGISTRY_FILE, registry)
    except Exception as exc:
        count = len(previous.get("rows", []))
        previous["startup_ok"] = False
        previous["startup_message"] = f"Hitelesítő tábla nem frissült indításkor, a korábbi helyi tábla maradt használatban ({count} prefix). Hiba: {exc}"
        return write_json_file(BANK_REGISTRY_FILE, previous)


def account_payload_from_body(body: dict[str, Any], registry: dict[str, Any]) -> dict[str, Any]:
    country = clean_cell(body.get("bank_country") or "HU").upper()
    if country != "HU":
        raise ValueError("Külföldi bankszámlákat egyelőre nem kezelünk.")
    currency = clean_cell(body.get("currency")).upper()
    if not currency:
        raise ValueError("A deviza megadása kötelező.")
    if currency not in CURRENCIES:
        raise ValueError(f"Ismeretlen deviza: {currency}")
    validation = validate_account_input(clean_cell(body.get("account_number")), registry)
    if not validation["ok"]:
        raise ValueError("\n".join(validation["errors"]))
    registry_bank = validation.get("bank") or {}
    return {
        "id": clean_cell(body.get("id")) or uuid.uuid4().hex,
        "company_id": ensure_company_id(clean_cell(body.get("company_id"))),
        "bank_country": "HU",
        "bank_name": clean_cell(body.get("bank_name")) or clean_cell(registry_bank.get("bank_name")),
        "currency": currency,
        "account_number": validation["formatted"],
        "account_digits": validation["normalized"],
        "prefix": validation.get("prefix", ""),
        "iban": validation.get("iban", ""),
        "input_type": validation.get("input_type", "HU"),
        "valid": True,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def account_identity(account: dict[str, Any]) -> str:
    stored = clean_cell(account.get("account_digits"))
    if stored:
        return stored
    digits = digits_only(account.get("account_number", ""))
    if len(digits) == 16:
        return digits + "0" * 8
    return digits


def find_duplicate_account(accounts: list[dict[str, Any]], payload: dict[str, Any]) -> dict[str, Any] | None:
    payload_id = clean_cell(payload.get("id"))
    payload_digits = account_identity(payload)
    payload_company = clean_cell(payload.get("company_id") or "default")
    if not payload_digits:
        return None
    for account in accounts:
        if clean_cell(account.get("id")) == payload_id:
            continue
        if clean_cell(account.get("company_id") or "default") != payload_company:
            continue
        if account_identity(account) == payload_digits:
            return account
    return None


def cleanup_saved_accounts(registry: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    registry = registry or load_bank_registry()
    accounts = load_accounts_file()
    kept: list[dict[str, Any]] = []
    removed: list[dict[str, Any]] = []
    for account in accounts:
        raw_value = clean_cell(account.get("iban")) or clean_cell(account.get("account_number")) or clean_cell(account.get("account_digits"))
        validation = validate_account_input(raw_value, registry)
        if validation["ok"]:
            kept.append(account)
        else:
            removed.append({**account, "validation_errors": validation["errors"]})
    if removed:
        save_accounts_file(kept)
    return removed


def read_json_body(handler: BaseHTTPRequestHandler) -> dict[str, Any]:
    length = int(handler.headers.get("Content-Length", "0") or "0")
    payload = handler.rfile.read(length).decode("utf-8") if length else "{}"
    data = json.loads(payload or "{}")
    if not isinstance(data, dict):
        raise ValueError("Hibás kérés.")
    return data


def import_accounts_from_upload(data: bytes, filename: str, company_id: str | None = None) -> dict[str, Any]:
    _, _, headers, rows = workbook_rows(data, filename)
    if not headers:
        raise ValueError("Nem található fejlécsor.")
    header_map = {loose_key(header): header for header in headers}
    def pick(*names: str) -> str:
        for name in names:
            found = header_map.get(loose_key(name))
            if found:
                return found
        return ""
    country_col = pick("Bank országa", "Ország", "Country")
    bank_col = pick("Bank neve", "Bank", "Bank name")
    account_col = pick("Számlaszám", "Bankszámlaszám", "Account", "IBAN")
    currency_col = pick("Deviza", "Pénznem", "Currency")
    if not account_col:
        raise ValueError("Az import fájlban kell egy Számlaszám vagy IBAN oszlop.")
    if not currency_col:
        raise ValueError("Az import fájlban kell egy Deviza oszlop.")
    registry = load_bank_registry()
    existing = load_accounts_file()
    by_id = {row.get("id"): row for row in existing}
    cid = ensure_company_id(company_id)
    by_digits = {account_identity(row): row for row in existing if account_identity(row) and clean_cell(row.get("company_id") or "default") == cid}
    added = 0
    errors: list[str] = []
    for line_no, raw_row in enumerate(rows, start=2):
        if row_is_empty(list(raw_row)):
            continue
        row = row_dict(headers, list(raw_row))
        try:
            payload = account_payload_from_body({
                "bank_country": row.get(country_col) or "HU",
                "company_id": cid,
                "bank_name": row.get(bank_col) or "",
                "currency": row.get(currency_col) or "",
                "account_number": row.get(account_col) or "",
            }, registry)
            duplicate = by_digits.get(account_identity(payload))
            if duplicate:
                errors.append(f"{line_no}. sor: ez a bankszámla már szerepel: {payload['account_number']}")
                continue
            by_id[payload["id"]] = payload
            by_digits[account_identity(payload)] = payload
            added += 1
        except ValueError as exc:
            errors.append(f"{line_no}. sor: {str(exc).splitlines()[0]}")
    save_accounts_file(list(by_id.values()))
    all_rows = list(by_id.values())
    return {"added": added, "errors": errors, "accounts": [row for row in all_rows if clean_cell(row.get("company_id") or "default") == cid]}


def build_accounts_template_xlsx() -> bytes:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Bankszamlak"
    headers = ["Bank országa", "Bank neve", "Számlaszám", "Deviza"]
    sheet.append(headers)
    sheet.append(["HU", "Minta Bank", "11111111-22222222-33333333", "HUF"])
    header_fill = PatternFill("solid", fgColor="263449")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    sheet.column_dimensions["A"].width = 16
    sheet.column_dimensions["B"].width = 34
    sheet.column_dimensions["C"].width = 32
    sheet.column_dimensions["D"].width = 12
    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def normalize_bic(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", clean_cell(value).upper())


def validate_bic(value: str) -> bool:
    bic = normalize_bic(value)
    return bool(re.fullmatch(r"[A-Z]{4}[A-Z]{2}[A-Z0-9]{2}([A-Z0-9]{3})?", bic))


def lookup_bic_online(bic: str = "", iban: str = "") -> dict[str, Any]:
    iban = normalize_iban(iban)
    bic = normalize_bic(bic)
    if iban:
        url = f"https://openiban.com/validate/{iban}?getBIC=true&validateBankCode=true"
        try:
            request = urllib.request.Request(url, headers={"User-Agent": "Banki import konvertalo"})
            with urllib.request.urlopen(request, timeout=10) as response:
                data = json.loads(response.read().decode("utf-8"))
            bank = data.get("bankData") or {}
            if bank:
                city = clean_cell(bank.get("city"))
                zip_code = clean_cell(bank.get("zip"))
                return {
                    "found": True,
                    "bic": clean_cell(bank.get("bic")) or bic,
                    "bank_name": clean_cell(bank.get("name")),
                    "bank_address": " ".join(part for part in [zip_code, city] if part),
                    "source": "openiban",
                    "source_message": "OpenIBAN alapján kitöltve. Az OpenIBAN BIC adat csak bizonyos országokra érhető el.",
                }
        except Exception:
            pass
    if bic:
        return {"found": False, "error": "A teljes BIC név/cím adatbázis hivatalosan SWIFTRef/licences forrás. Ehhez most nincs beállított API-kulcs; add meg kézzel, vagy IBAN alapján próbáld."}
    return {"found": False, "error": "Adj meg IBAN-t vagy SWIFT/BIC kódot."}


def partner_payload_from_body(body: dict[str, Any], registry: dict[str, Any]) -> dict[str, Any]:
    name = clean_cell(body.get("name"))
    if not name:
        raise ValueError("A partner neve kötelező.")
    account_number = clean_cell(body.get("account_number"))
    iban = clean_cell(body.get("iban"))
    swift = normalize_bic(body.get("swift_bic", ""))
    country = clean_cell(body.get("country") or "HU").upper()
    bank_name = clean_cell(body.get("bank_name"))
    bank_address = clean_cell(body.get("bank_address"))
    normalized_account = ""
    if account_number:
        validation = validate_hu_account(account_number, registry)
        if not validation["ok"]:
            raise ValueError("\n".join(validation["errors"]))
        normalized_account = validation["normalized"]
        account_number = validation["formatted"]
        bank = validation.get("bank") or {}
        bank_name = bank_name or clean_cell(bank.get("bank_name"))
        country = "HU"
    if iban:
        iban_result = validate_iban_syntax(iban)
        if not iban_result["ok"]:
            raise ValueError("\n".join(iban_result["errors"]))
        iban = iban_result["iban"]
        if iban.startswith("HU") and not account_number:
            validation = validate_hu_account(iban[4:], registry)
            if validation["ok"]:
                normalized_account = validation["normalized"]
                account_number = validation["formatted"]
                bank = validation.get("bank") or {}
                bank_name = bank_name or clean_cell(bank.get("bank_name"))
    if swift and not validate_bic(swift):
        raise ValueError("A SWIFT/BIC 8 vagy 11 karakteres ISO 9362 formátum legyen.")
    if not account_number and not iban:
        raise ValueError("Adj meg magyar számlaszámot vagy IBAN-t.")
    return {
        "id": clean_cell(body.get("id")) or uuid.uuid4().hex,
        "company_id": ensure_company_id(clean_cell(body.get("company_id"))),
        "partner_code": clean_cell(body.get("partner_code")),
        "name": name,
        "account_number": account_number,
        "account_digits": normalized_account,
        "iban": iban,
        "swift_bic": swift,
        "country": country,
        "address": clean_cell(body.get("address")),
        "bank_name": bank_name,
        "bank_address": bank_address,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def partner_identity(partner: dict[str, Any]) -> str:
    return clean_cell(partner.get("partner_code")) or clean_cell(partner.get("iban")) or clean_cell(partner.get("account_digits")) or clean_cell(partner.get("account_number"))


def import_partners_from_upload(data: bytes, filename: str, company_id: str | None = None) -> dict[str, Any]:
    _, _, headers, rows = workbook_rows(data, filename)
    if not headers:
        raise ValueError("Nem található fejlécsor.")
    header_map = {loose_key(header): header for header in headers}
    def pick(*names: str) -> str:
        for name in names:
            found = header_map.get(loose_key(name))
            if found:
                return found
        return ""
    cols = {
        "partner_code": pick("Partner kód", "Partner kod", "Code"),
        "name": pick("Név", "Nev", "Partner neve", "Name"),
        "account_number": pick("Számlaszám", "Szamlaszam", "Magyar számlaszám"),
        "iban": pick("IBAN"),
        "swift_bic": pick("SWIFT/BIC", "SWIFT", "BIC"),
        "country": pick("Ország", "Orszag", "Country"),
        "address": pick("Partner címe", "Partner cime", "Cím", "Address"),
        "bank_name": pick("Bank neve", "Bank name"),
        "bank_address": pick("Bank címe", "Bank cime", "Bank address"),
    }
    registry = load_bank_registry()
    existing = load_partners_file()
    cid = ensure_company_id(company_id)
    by_id = {row.get("id"): row for row in existing}
    identities = {partner_identity(row): row for row in existing if partner_identity(row) and clean_cell(row.get("company_id") or "default") == cid}
    added = 0
    errors: list[str] = []
    for line_no, raw_row in enumerate(rows, start=2):
        if row_is_empty(list(raw_row)):
            continue
        row = row_dict(headers, list(raw_row))
        try:
            payload = partner_payload_from_body({**{key: row.get(col, "") if col else "" for key, col in cols.items()}, "company_id": cid}, registry)
            ident = partner_identity(payload)
            if ident and ident in identities:
                errors.append(f"{line_no}. sor: ez a partner már szerepel: {ident}")
                continue
            by_id[payload["id"]] = payload
            identities[ident] = payload
            added += 1
        except ValueError as exc:
            errors.append(f"{line_no}. sor: {str(exc).splitlines()[0]}")
    all_rows = list(by_id.values())
    save_partners_file(all_rows)
    return {"added": added, "errors": errors, "partners": [row for row in all_rows if clean_cell(row.get("company_id") or "default") == cid]}


def build_partners_template_xlsx() -> bytes:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Partnerek"
    headers = ["Partner kód", "Név", "Számlaszám", "IBAN", "SWIFT/BIC", "Ország", "Partner címe", "Bank neve", "Bank címe"]
    sheet.append(headers)
    sheet.append(["P001", "Minta Partner Kft", "44444444-55555555-66666666", "", "", "HU", "Minta cím", "", ""])
    sheet.append(["P002", "Minta Partner GmbH", "", "AT451100000421840000", "BKAUATWW", "AT", "Minta cím", "", ""])
    header_fill = PatternFill("solid", fgColor="263449")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    widths = [16, 32, 32, 34, 16, 10, 34, 34, 34]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[openpyxl.utils.get_column_letter(index)].width = width
    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def default_settings() -> dict[str, Any]:
    return {
        "active_company_id": "default",
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
        if base.get("active_bank") not in BANKS:
            base["active_bank"] = "erste"
        if base.get("active_format") not in FORMATS:
            base["active_format"] = next((fmt_id for fmt_id, fmt in FORMATS.items() if fmt.get("bank") == base.get("active_bank")), "erste_huf_payord")
        if not any(company.get("id") == base.get("active_company_id") for company in load_companies_file()):
            base["active_company_id"] = "default"
        return base
    except Exception:
        return default_settings()


def save_settings_file(data: dict[str, Any]) -> dict[str, Any]:
    cleaned = default_settings()
    cleaned.update(data if isinstance(data, dict) else {})
    if cleaned.get("active_bank") not in BANKS:
        cleaned["active_bank"] = "erste"
    if cleaned.get("active_format") not in FORMATS:
        cleaned["active_format"] = next((fmt_id for fmt_id, fmt in FORMATS.items() if fmt.get("bank") == cleaned.get("active_bank")), "erste_huf_payord")
    if not any(company.get("id") == cleaned.get("active_company_id") for company in load_companies_file()):
        cleaned["active_company_id"] = ensure_company_id("default")
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
        # PWA root-scope assets served from static/
        if path in ("/manifest.webmanifest", "/sw.js"):
            fname = path.lstrip("/")
            fp = Path(__file__).parent / "static" / fname
            if fp.is_file():
                ctype = "application/manifest+json" if path.endswith(".webmanifest") else "application/javascript; charset=utf-8"
                data = fp.read_bytes()
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", ctype)
                self.send_header("Content-Length", str(len(data)))
                if path == "/sw.js":
                    self.send_header("Service-Worker-Allowed", "/")
                self.end_headers()
                self.wfile.write(data); return
            self.send_error(HTTPStatus.NOT_FOUND); return
        if path.startswith("/static/"):
            rel = path[len("/static/"):]
            # Allow only safe relative paths under static/ (optionally one subdir like js/)
            if ".." in rel or rel.startswith("/") or rel.startswith("\\") or rel.startswith("."):
                self.send_error(HTTPStatus.NOT_FOUND); return
            parts = rel.split("/")
            if len(parts) > 3 or any(not p or p.startswith(".") for p in parts):
                self.send_error(HTTPStatus.NOT_FOUND); return
            fp = (Path(__file__).parent / "static").joinpath(*parts)
            if not fp.is_file():
                self.send_error(HTTPStatus.NOT_FOUND); return
            ctype = "text/css; charset=utf-8" if rel.endswith(".css") else (
                "application/javascript; charset=utf-8" if rel.endswith(".js")
                else "image/svg+xml" if rel.endswith(".svg")
                else "application/octet-stream"
            )
            data = fp.read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", ctype)
            self.send_header("Cache-Control", "public, max-age=300")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if path in {"/", "/index.html"}:
            body = (
                HTML_PAGE
                .replace("__BANKS_JSON__", json.dumps(BANKS, ensure_ascii=False))
                .replace("__FORMATS_JSON__", json.dumps(FORMATS, ensure_ascii=False))
                .replace("__FIELDS_JSON__", json.dumps(FIELDS, ensure_ascii=False))
                .replace("__CURRENCIES_JSON__", json.dumps(CURRENCIES, ensure_ascii=False))
            )
            text_response(self, body)
            return
        if path == "/api/settings":
            json_response(self, load_settings_file())
            return
        if path == "/api/companies":
            settings = load_settings_file()
            json_response(self, {"companies": load_companies_file(), "active_company_id": settings.get("active_company_id", "default")})
            return
        if path == "/api/accounts":
            params = parse_qs(parsed.query)
            company_id = ensure_company_id((params.get("company_id") or [""])[0])
            removed = cleanup_saved_accounts()
            json_response(self, {"accounts": accounts_for_company(company_id), "removed_invalid": len(removed), "company_id": company_id})
            return
        if path == "/api/partners":
            params = parse_qs(parsed.query)
            company_id = ensure_company_id((params.get("company_id") or [""])[0])
            json_response(self, {"partners": partners_for_company(company_id), "company_id": company_id})
            return
        if path == "/api/bic-lookup":
            params = parse_qs(parsed.query)
            json_response(self, lookup_bic_online((params.get("bic") or [""])[0], (params.get("iban") or [""])[0]))
            return
        if path == "/api/audit":
            try: limit = int((parse_qs(urlparse(self.path).query).get("limit") or ["200"])[0])
            except Exception: limit = 200
            json_response(self, {"entries": read_audit(min(max(limit, 1), 1000))})
            return
        if path == "/healthz":
            json_response(self, {"ok": True, "ts": int(_time_mod.time() * 1000)})
            return
        if path == "/api/backup":
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                for fp in [SETTINGS_FILE, ACCOUNTS_FILE, COMPANIES_FILE, PARTNERS_FILE,
                           TRASH_FILE, AUDIT_FILE, SCHEMA_VERSION_FILE, BANK_REGISTRY_FILE]:
                    if fp and fp.is_file():
                        try:
                            zf.write(fp, arcname=fp.name)
                        except Exception as exc:
                            log.warning("backup: skip %s (%s)", fp.name, exc)
                meta = {
                    "created_at": datetime.now().isoformat(timespec="seconds"),
                    "schema_version": CURRENT_SCHEMA_VERSION,
                    "files": sorted([n for n in zf.namelist()]),
                }
                zf.writestr("_backup_manifest.json", json.dumps(meta, ensure_ascii=False, indent=2))
            data = buf.getvalue()
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/zip")
            self.send_header("Content-Disposition", f"attachment; filename=banki_backup_{stamp}.zip")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            try: audit_log("backup.download", "", f"bytes={len(data)} files={len(meta['files'])}")
            except Exception: pass
            return
        if path == "/api/bank-registry":
            registry = load_bank_registry()
            json_response(self, bank_registry_response(registry))
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
        if path == "/accounts-template.xlsx":
            data = build_accounts_template_xlsx()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", "attachment; filename=bankszamla_sablon.xlsx")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if path == "/partners-template.xlsx":
            data = build_partners_template_xlsx()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", "attachment; filename=partner_sablon.xlsx")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        try:
            if self.path == "/api/settings":
                json_response(self, save_settings_file(read_json_body(self)))
                return
            parsed_post = urlparse(self.path)
            post_path = parsed_post.path
            post_params = parse_qs(parsed_post.query)
            if post_path == "/api/companies":
                body = read_json_body(self)
                name = clean_cell(body.get("name"))
                if not name:
                    raise ValueError("A cég neve kötelező.")
                company_id = clean_cell(body.get("id"))
                companies = load_companies_file()
                company = None
                if company_id:
                    for row in companies:
                        if row.get("id") == company_id:
                            row["name"] = name
                            company = row
                            break
                if company is None:
                    company = {"id": uuid.uuid4().hex, "name": name}
                    companies.append(company)
                save_companies_file(companies); audit_log('company.save', company.get('id',''), company.get('name',''))
                settings = load_settings_file()
                settings["active_company_id"] = company["id"]
                save_settings_file(settings)
                json_response(self, {"company": company, "companies": companies})
                return
            if post_path == "/api/accounts":
                body = read_json_body(self)
                company_id = ensure_company_id((post_params.get("company_id") or [body.get("company_id") or ""])[0])
                body["company_id"] = company_id
                registry = load_bank_registry()
                payload = account_payload_from_body(body, registry)
                accounts = load_accounts_file()
                duplicate = find_duplicate_account(accounts, payload)
                if duplicate:
                    raise ValueError(f"Ez a bankszámla már szerepel a listában: {duplicate.get('account_number') or payload['account_number']}")
                found = False
                for index, account in enumerate(accounts):
                    if account.get("id") == payload["id"]:
                        accounts[index] = payload
                        found = True
                        break
                if not found:
                    accounts.append(payload)
                save_accounts_file(accounts)
                json_response(self, {"account": payload, "accounts": accounts_for_company(company_id)})
                return
            if post_path == "/api/accounts/delete":
                body = read_json_body(self)
                company_id = ensure_company_id((post_params.get("company_id") or [body.get("company_id") or ""])[0])
                account_id = clean_cell(body.get("id"))
                all_acc = load_accounts_file()
                victim = next((row for row in all_acc if row.get("id") == account_id and clean_cell(row.get("company_id") or "default") == company_id), None)
                if victim: push_to_trash("accounts", victim)
                accounts = [row for row in all_acc if not (row.get("id") == account_id and clean_cell(row.get("company_id") or "default") == company_id)]
                save_accounts_file(accounts)
                audit_log("account.delete", company_id, str(victim.get("account_number") if victim else account_id))
                json_response(self, {"accounts": accounts_for_company(company_id)})
                return
            if post_path == "/api/accounts/import":
                form = get_upload(self)
                filename, data = upload_bytes(form)
                company_id = ensure_company_id((post_params.get("company_id") or [""])[0])
                json_response(self, import_accounts_from_upload(data, filename, company_id))
                return
            if post_path == "/api/partners":
                body = read_json_body(self)
                company_id = ensure_company_id((post_params.get("company_id") or [body.get("company_id") or ""])[0])
                body["company_id"] = company_id
                registry = load_bank_registry()
                payload = partner_payload_from_body(body, registry)
                partners = load_partners_file()
                ident = partner_identity(payload)
                for partner in partners:
                    if (
                        partner.get("id") != payload["id"]
                        and ident
                        and partner_identity(partner) == ident
                        and clean_cell(partner.get("company_id") or "default") == company_id
                    ):
                        raise ValueError(f"Ez a partner már szerepel: {ident}")
                found = False
                for index, partner in enumerate(partners):
                    if partner.get("id") == payload["id"] and clean_cell(partner.get("company_id") or "default") == company_id:
                        partners[index] = payload
                        found = True
                        break
                if not found:
                    partners.append(payload)
                save_partners_file(partners); audit_log('partner.save', company_id, payload.get('name',''))
                json_response(self, {"partner": payload, "partners": partners_for_company(company_id)})
                return
            if post_path == "/api/partners/delete":
                body = read_json_body(self)
                company_id = ensure_company_id((post_params.get("company_id") or [body.get("company_id") or ""])[0])
                partner_id = clean_cell(body.get("id"))
                all_partners = load_partners_file()
                victim = next((row for row in all_partners if row.get("id") == partner_id and clean_cell(row.get("company_id") or "default") == company_id), None)
                if victim: push_to_trash("partners", victim)
                partners = [row for row in all_partners if not (row.get("id") == partner_id and clean_cell(row.get("company_id") or "default") == company_id)]
                save_partners_file(partners)
                audit_log("partner.delete", company_id, str(victim.get("name") if victim else partner_id))
                json_response(self, {"partners": partners_for_company(company_id)})
                return
            if post_path == "/api/partners/import":
                form = get_upload(self)
                filename, data = upload_bytes(form)
                company_id = ensure_company_id((post_params.get("company_id") or [""])[0])
                json_response(self, import_partners_from_upload(data, filename, company_id))
                return
            if post_path == "/api/bank-registry":
                json_response(self, save_bank_registry(read_json_body(self)))
                return
            if post_path == "/api/bank-registry/refresh":
                try:
                    read_json_body(self)
                except Exception:
                    pass
                try:
                    registry = refresh_bank_registry()
                    json_response(self, bank_registry_response(registry, status="ok"))
                except Exception as exc:
                    registry = load_bank_registry()
                    json_response(self, bank_registry_response(registry, status="stale", error=str(exc)))
                return
            if post_path == "/api/inspect":
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
            if post_path == "/api/convert":
                form = get_upload(self)
                filename, data = upload_bytes(form)
                config = json.loads(field_value(form, "config", "{}") or "{}")
                text, count = convert_records(data, filename, config)
                encoding = config.get("encoding") or "cp1250"
                codec = "cp852" if encoding == "cp852" else ("cp1250" if encoding == "cp1250" else "utf-8")
                payload = text.encode(codec, errors="replace")
                self.send_response(HTTPStatus.OK)
                charset = "ibm852" if codec == "cp852" else ("windows-1250" if codec == "cp1250" else "utf-8")
                self.send_header("Content-Type", f"text/plain; charset={charset}")
                self.send_header("Content-Disposition", "attachment; filename=payord_import.txt")
                self.send_header("X-Filename", "payord_import.txt")
                self.send_header("X-Record-Count", str(count))
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            if post_path == "/api/partners/restore":
                body = read_json_body(self)
                row = pop_from_trash("partners", clean_cell(body.get("id")), ensure_company_id(body.get("company_id") or ""))
                if row:
                    partners = load_partners_file(); partners.append(row); save_partners_file(partners)
                    audit_log("partner.restore", row.get("company_id",""), str(row.get("name","")))
                json_response(self, {"restored": bool(row)})
                return
            if post_path == "/api/accounts/restore":
                body = read_json_body(self)
                row = pop_from_trash("accounts", clean_cell(body.get("id")), ensure_company_id(body.get("company_id") or ""))
                if row:
                    accounts = load_accounts_file(); accounts.append(row); save_accounts_file(accounts)
                    audit_log("account.restore", row.get("company_id",""), str(row.get("account_number","")))
                json_response(self, {"restored": bool(row)})
                return
            if post_path == "/api/companies/delete":
                body = read_json_body(self)
                cid = clean_cell(body.get("id"))
                all_co = load_companies_file()
                victim = next((c for c in all_co if c.get("id") == cid), None)
                if victim: push_to_trash("companies", victim)
                companies = [c for c in all_co if c.get("id") != cid]
                save_companies_file(companies)
                audit_log("company.delete", cid, str(victim.get("name") if victim else cid))
                json_response(self, {"companies": companies})
                return
            if post_path == "/api/companies/restore":
                body = read_json_body(self)
                row = pop_from_trash("companies", clean_cell(body.get("id")))
                if row:
                    companies = load_companies_file(); companies.append(row); save_companies_file(companies)
                    audit_log("company.restore", row.get("id",""), str(row.get("name","")))
                json_response(self, {"restored": bool(row)})
                return
            self.send_error(HTTPStatus.NOT_FOUND)
        except Exception as exc:
            message = str(exc)
            json_response(self, {"error": message, "errors": [line for line in message.splitlines() if line]}, status=400)

    def log_message(self, fmt: str, *args: Any) -> None:
        log.info("%s - %s", self.address_string(), fmt % args)



# ---------- Iteration 6 additions: trash, audit, bulk, healthz ----------

import threading
import time as _time_mod

def _load_trash() -> dict:
    return load_json_file(TRASH_FILE, {"partners": [], "accounts": [], "companies": []})

def _save_trash(data: dict) -> dict:
    return write_json_file(TRASH_FILE, data)

def push_to_trash(kind: str, item: dict) -> None:
    t = _load_trash()
    t.setdefault(kind, [])
    item = dict(item); item["_trashed_at"] = _time_mod.time()
    t[kind].append(item)
    # keep last 200
    t[kind] = t[kind][-200:]
    _save_trash(t)

def pop_from_trash(kind: str, item_id: str, company_id: str = "") -> dict | None:
    t = _load_trash()
    rows = t.get(kind, [])
    for i, row in enumerate(rows):
        if row.get("id") == item_id and (not company_id or clean_cell(row.get("company_id") or "default") == company_id):
            rows.pop(i)
            t[kind] = rows
            _save_trash(t)
            row.pop("_trashed_at", None)
            return row
    return None

def audit_log(action: str, company_id: str = "", detail: str = "") -> None:
    try:
        line = json.dumps({
            "ts": _time_mod.time() * 1000,
            "action": action,
            "company_id": company_id,
            "detail": detail[:500],
        }, ensure_ascii=False)
        with open(AUDIT_FILE, "a", encoding="utf-8") as fh:
            fh.write(line + "\n")
        _rotate_audit_if_needed()
    except Exception:
        log.exception("audit_log failed")

def _rotate_audit_if_needed() -> None:
    """Keep audit log under AUDIT_MAX_LINES; rotate older entries to .1 backup."""
    try:
        if not AUDIT_FILE.exists():
            return
        with open(AUDIT_FILE, "r", encoding="utf-8") as fh:
            lines = fh.readlines()
        if len(lines) <= AUDIT_MAX_LINES:
            return
        keep = lines[-AUDIT_MAX_LINES:]
        archive = lines[:-AUDIT_MAX_LINES]
        backup = AUDIT_FILE.with_suffix(AUDIT_FILE.suffix + ".1")
        with open(backup, "a", encoding="utf-8") as fh:
            fh.writelines(archive)
        with open(AUDIT_FILE, "w", encoding="utf-8") as fh:
            fh.writelines(keep)
        log.info("audit log rotated: kept %d lines, archived %d to %s",
                 len(keep), len(archive), backup.name)
    except Exception:
        log.exception("audit rotation failed")

# ---------- Schema version + migration ----------
def _read_schema_version() -> int:
    if not SCHEMA_VERSION_FILE.exists():
        return 0
    try:
        return int(json.loads(SCHEMA_VERSION_FILE.read_text(encoding="utf-8")).get("version", 0))
    except Exception:
        return 0

def _write_schema_version(v: int) -> None:
    SCHEMA_VERSION_FILE.write_text(
        json.dumps({"version": v, "updated_at": _time_mod.time()}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

def run_schema_migrations() -> None:
    """Run idempotent migrations. Add new ones as elif blocks bumping CURRENT_SCHEMA_VERSION."""
    v = _read_schema_version()
    if v >= CURRENT_SCHEMA_VERSION:
        return
    log.info("running schema migrations: %d -> %d", v, CURRENT_SCHEMA_VERSION)
    if v < 1:
        # v0 -> v1: normalize company_id on accounts & partners (legacy by_company shape)
        try:
            save_accounts_file(load_accounts_file())
            save_partners_file(load_partners_file())
            save_companies_file(load_companies_file())
            log.info("migration v1: normalized accounts/partners/companies")
        except Exception:
            log.exception("migration v1 failed")
    if v < 2:
        # v1 -> v2: ensure trash file exists with all kinds
        try:
            t = _load_trash()
            for k in ("partners", "accounts", "companies"):
                t.setdefault(k, [])
            _save_trash(t)
            log.info("migration v2: trash file initialized")
        except Exception:
            log.exception("migration v2 failed")
    _write_schema_version(CURRENT_SCHEMA_VERSION)
    log.info("schema migrations complete: now at v%d", CURRENT_SCHEMA_VERSION)

def read_audit(limit: int = 200) -> list[dict]:
    if not AUDIT_FILE.exists():
        return []
    try:
        with open(AUDIT_FILE, "r", encoding="utf-8") as fh:
            lines = fh.readlines()
        out = []
        for ln in lines[-limit:]:
            ln = ln.strip()
            if not ln: continue
            try: out.append(json.loads(ln))
            except Exception: pass
        return out
    except Exception:
        return []

def background_registry_refresh(interval_sec: int | None = None) -> None:
    if interval_sec is None:
        interval_sec = int(MNB_REFRESH_HOURS * 3600)
    def loop():
        while True:
            try:
                _time_mod.sleep(interval_sec)
                refresh_bank_registry()
            except Exception:
                pass
    t = threading.Thread(target=loop, daemon=True, name="banki-mnb-refresh")
    t.start()


def main() -> None:
    run_schema_migrations()
    registry = refresh_bank_registry_at_startup()
    background_registry_refresh()
    log.info(registry.get("startup_message") or "Hitelesítő tábla állapota betöltve.")
    save_companies_file(load_companies_file())
    save_accounts_file(load_accounts_file())
    removed = cleanup_saved_accounts(registry)
    if removed:
        log.info("%d korábban mentett, érvénytelen saját bankszámla eltávolítva.", len(removed))
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    log.info("PAYORD konverter fut: http://%s:%s", HOST, PORT)
    server.serve_forever()


if __name__ == "__main__":
    main()
