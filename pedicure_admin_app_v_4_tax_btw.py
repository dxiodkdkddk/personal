#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pedicure Administratie Applicatie (Tkinter + SQLite)

v4 – Wat zit erin:
- Agenda met afspraken (tkcalendar)
- Cliëntenbeheer (naam, e-mail, taal nl/fr/en/ar)
- Prijslijst met manipulaties & prijzen
- Reçus genereren (PDF) + mailen in taal van de klant
- Overzichten: dag/week/maand/jaar en custom periode
- Export CSV/Excel (xlsxwriter)
- Firmanaam in bestandsnamen
- Totale inkomsten vandaag/week/maand/jaar
- **Btw-voet instelbaar** via menu (default 21%)
- **Reçu met btw-uitsplitsing** (Netto / Btw / Totaal)
- **Belastingdocument huidig jaar** met maandsommen + netto/btw/totaal

Benodigdheden:
    pip install tkcalendar reportlab babel python-dotenv xlsxwriter
"""

import os
import sqlite3
import datetime as dt
from dataclasses import dataclass
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

# Externe libs (optioneel)
try:
    from tkcalendar import Calendar, DateEntry  # type: ignore
except Exception:
    Calendar = None
    DateEntry = None

try:
    from reportlab.lib.pagesizes import A4  # type: ignore
    from reportlab.pdfgen import canvas as pdfcanvas  # type: ignore
    from reportlab.lib.units import mm  # type: ignore
    from reportlab.lib import colors  # type: ignore
except Exception:
    pdfcanvas = None

try:
    from dotenv import load_dotenv  # type: ignore
    load_dotenv()
except Exception:
    pass

APP_DIR = Path.home() / ".pedicure_app"
DB_PATH = APP_DIR / "pedicure.db"
PDF_DIR = APP_DIR / "receipts"
APP_DIR.mkdir(parents=True, exist_ok=True)
PDF_DIR.mkdir(parents=True, exist_ok=True)

SUPPORTED_LANGS = ["nl", "fr", "en", "ar"]  # ar ~ Algerijns (Arabisch)

# ---- Vertalingen ------------------------------------------------------------
T = {
    "nl": {
        "app_title": "Pedicure Administratie",
        "company": "Firma",
        "admin": "Beheerder",
        "language": "Taal",
        "settings": "Instellingen",
        "set_vat": "Stel btw (%) in",
        "vat": "Btw",
        "dashboard": "Overzicht",
        "agenda": "Agenda",
        "clients": "Cliënten",
        "pricelist": "Prijslijst",
        "receipts": "Reçus",
        "totals_today": "Totaal vandaag",
        "totals_week": "Totaal deze week",
        "totals_month": "Totaal deze maand",
        "totals_year": "Totaal dit jaar",
        "add": "Toevoegen",
        "edit": "Wijzigen",
        "delete": "Verwijderen",
        "save": "Opslaan",
        "cancel": "Annuleren",
        "name": "Naam",
        "email": "E-mail",
        "phone": "Telefoon",
        "notes": "Notities",
        "client_lang": "Klanttaal",
        "manipulation": "Manipulatie",
        "price": "Prijs",
        "change_manip": "Verander manipulatie",
        "change_price": "Verander prijs",
        "change_client": "Verander cliënt",
        "print_receipt": "Druk reçu",
        "print_day": "Druk dag reçu",
        "print_week": "Druk week reçu",
        "print_month": "Druk maand reçu",
        "print_year": "Druk jaar reçu",
        "email_receipt": "E-mail reçu",
        "new_appointment": "Nieuwe afspraak",
        "date": "Datum",
        "time": "Tijd",
        "duration_min": "Duur (min)",
        "select_client": "Selecteer cliënt",
        "select_manips": "Selecteer manipulaties",
        "total": "Totaal",
        "first_run_title": "Eerste opstart",
        "enter_company": "Voer firmanaam in:",
        "enter_admin": "Voer administratornaam in:",
        "choose_lang": "Kies basistaal (nl/fr/en/ar):",
        "enter_manip_list": "Voeg manipulaties toe (naam en prijs)",
        "receipt": "Reçu",
        "no_calendar": "tkcalendar niet geïnstalleerd. Installeer met: pip install tkcalendar",
        "no_pdf": "reportlab niet geïnstalleerd. Installeer met: pip install reportlab",
        "confirm": "Bevestigen",
        "are_you_sure": "Ben je zeker?",
        "all_clients": "Alle cliënten",
        "filter_client": "Filter cliënt",
        "print_period": "Druk periode",
        "start_date": "Van",
        "end_date": "Tot",
        "export_csv": "Exporteer CSV",
        "export_excel": "Exporteer Excel",
        "subtotal_by_manip": "Subtotaal per manipulatie",
        "tax_doc": "Belastingdocument huidig jaar",
    },
    "fr": {
        "app_title": "Administration Pédicure",
        "company": "Société",
        "admin": "Administrateur",
        "language": "Langue",
        "settings": "Paramètres",
        "set_vat": "Définir TVA (%)",
        "vat": "TVA",
        "dashboard": "Tableau de bord",
        "agenda": "Agenda",
        "clients": "Clients",
        "pricelist": "Tarifs",
        "receipts": "Reçus",
        "totals_today": "Total aujourd'hui",
        "totals_week": "Total cette semaine",
        "totals_month": "Total ce mois",
        "totals_year": "Total cette année",
        "add": "Ajouter",
        "edit": "Modifier",
        "delete": "Supprimer",
        "save": "Enregistrer",
        "cancel": "Annuler",
        "name": "Nom",
        "email": "E-mail",
        "phone": "Téléphone",
        "notes": "Notes",
        "client_lang": "Langue client",
        "manipulation": "Manipulation",
        "price": "Prix",
        "change_manip": "Changer manipulation",
        "change_price": "Changer prix",
        "change_client": "Changer client",
        "print_receipt": "Imprimer reçu",
        "print_day": "Imprimer reçu du jour",
        "print_week": "Imprimer reçu hebdo",
        "print_month": "Imprimer reçu mensuel",
        "print_year": "Imprimer reçu annuel",
        "email_receipt": "Envoyer reçu",
        "new_appointment": "Nouveau rendez-vous",
        "date": "Date",
        "time": "Heure",
        "duration_min": "Durée (min)",
        "select_client": "Sélectionner client",
        "select_manips": "Sélectionner manipulations",
        "total": "Total",
        "first_run_title": "Première ouverture",
        "enter_company": "Entrez le nom de la société :",
        "enter_admin": "Entrez le nom de l'administrateur :",
        "choose_lang": "Choisissez la langue (nl/fr/en/ar):",
        "enter_manip_list": "Ajoutez des manipulations (nom et prix)",
        "receipt": "Reçu",
        "no_calendar": "tkcalendar non installé. Installez : pip install tkcalendar",
        "no_pdf": "reportlab non installé. Installez : pip install reportlab",
        "confirm": "Confirmer",
        "are_you_sure": "Êtes-vous sûr ?",
        "all_clients": "Tous les clients",
        "filter_client": "Filtrer client",
        "print_period": "Imprimer période",
        "start_date": "Du",
        "end_date": "Au",
        "export_csv": "Exporter CSV",
        "export_excel": "Exporter Excel",
        "subtotal_by_manip": "Sous-total par manipulation",
        "tax_doc": "Document fiscal (année en cours)",
    },
    "en": {
        "app_title": "Pedicure Admin",
        "company": "Company",
        "admin": "Administrator",
        "language": "Language",
        "settings": "Settings",
        "set_vat": "Set VAT (%)",
        "vat": "VAT",
        "dashboard": "Dashboard",
        "agenda": "Agenda",
        "clients": "Clients",
        "pricelist": "Price List",
        "receipts": "Receipts",
        "totals_today": "Total today",
        "totals_week": "Total this week",
        "totals_month": "Total this month",
        "totals_year": "Total this year",
        "add": "Add",
        "edit": "Edit",
        "delete": "Delete",
        "save": "Save",
        "cancel": "Cancel",
        "name": "Name",
        "email": "Email",
        "phone": "Phone",
        "notes": "Notes",
        "client_lang": "Client language",
        "manipulation": "Procedure",
        "price": "Price",
        "change_manip": "Change procedure",
        "change_price": "Change price",
        "change_client": "Change client",
        "print_receipt": "Print receipt",
        "print_day": "Print day receipt",
        "print_week": "Print week receipt",
        "print_month": "Print month receipt",
        "print_year": "Print year receipt",
        "email_receipt": "Email receipt",
        "new_appointment": "New appointment",
        "date": "Date",
        "time": "Time",
        "duration_min": "Duration (min)",
        "select_client": "Select client",
        "select_manips": "Select procedures",
        "total": "Total",
        "first_run_title": "First run",
        "enter_company": "Enter company name:",
        "enter_admin": "Enter administrator name:",
        "choose_lang": "Choose base language (nl/fr/en/ar):",
        "enter_manip_list": "Add procedures (name and price)",
        "receipt": "Receipt",
        "no_calendar": "tkcalendar not installed. Install: pip install tkcalendar",
        "no_pdf": "reportlab not installed. Install: pip install reportlab",
        "confirm": "Confirm",
        "are_you_sure": "Are you sure?",
        "all_clients": "All clients",
        "filter_client": "Filter client",
        "print_period": "Print period",
        "start_date": "From",
        "end_date": "To",
        "export_csv": "Export CSV",
        "export_excel": "Export Excel",
        "subtotal_by_manip": "Subtotal by procedure",
        "tax_doc": "Tax document (current year)",
    },
    "ar": {
        "app_title": "إدارة العناية بالقدم",
        "company": "الشركة",
        "admin": "المدير",
        "language": "اللغة",
        "settings": "الإعدادات",
        "set_vat": "تعيين الضريبة (%)",
        "vat": "الضريبة",
        "dashboard": "لوحة التحكم",
        "agenda": "الأجندة",
        "clients": "العملاء",
        "pricelist": "قائمة الأسعار",
        "receipts": "الإيصالات",
        "totals_today": "إجمالي اليوم",
        "totals_week": "إجمالي هذا الأسبوع",
        "totals_month": "إجمالي هذا الشهر",
        "totals_year": "إجمالي هذه السنة",
        "add": "إضافة",
        "edit": "تعديل",
        "delete": "حذف",
        "save": "حفظ",
        "cancel": "إلغاء",
        "name": "الاسم",
        "email": "البريد الإلكتروني",
        "phone": "الهاتف",
        "notes": "ملاحظات",
        "client_lang": "لغة العميل",
        "manipulation": "الإجراء",
        "price": "السعر",
        "change_manip": "تغيير الإجراء",
        "change_price": "تغيير السعر",
        "change_client": "تغيير العميل",
        "print_receipt": "طباعة إيصال",
        "print_day": "طباعة إيصال اليوم",
        "print_week": "طباعة إيصال الأسبوع",
        "print_month": "طباعة إيصال الشهر",
        "print_year": "طباعة إيصال السنة",
        "email_receipt": "إرسال الإيصال",
        "new_appointment": "موعد جديد",
        "date": "التاريخ",
        "time": "الوقت",
        "duration_min": "المدة (د)",
        "select_client": "اختر العميل",
        "select_manips": "اختر الإجراءات",
        "total": "الإجمالي",
        "first_run_title": "التشغيل الأول",
        "enter_company": "أدخل اسم الشركة:",
        "enter_admin": "أدخل اسم المدير:",
        "choose_lang": "اختر اللغة الأساسية (nl/fr/en/ar):",
        "enter_manip_list": "أضف إجراءات (الاسم والسعر)",
        "receipt": "إيصال",
        "no_calendar": "tkcalendar غير مثبت.",
        "no_pdf": "reportlab غير مثبت.",
        "confirm": "تأكيد",
        "are_you_sure": "هل أنت متأكد؟",
        "all_clients": "كل العملاء",
        "filter_client": "تصفية العميل",
        "print_period": "طباعة الفترة",
        "start_date": "من",
        "end_date": "إلى",
        "export_csv": "تصدير CSV",
        "export_excel": "تصدير Excel",
        "subtotal_by_manip": "الإجمالي حسب الإجراء",
        "tax_doc": "مستند الضرائب (هذه السنة)",
    },
}

# ---- Database ---------------------------------------------------------------
SCHEMA_SQL = """
PRAGMA foreign_keys = ON;
CREATE TABLE IF NOT EXISTS config (
    key TEXT PRIMARY KEY,
    value TEXT
);
CREATE TABLE IF NOT EXISTS clients (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    email TEXT,
    phone TEXT,
    notes TEXT,
    lang TEXT DEFAULT 'nl'
);
CREATE TABLE IF NOT EXISTS manipulations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    price_cents INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS appointments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER,
    date TEXT NOT NULL,   -- YYYY-MM-DD
    time TEXT NOT NULL,   -- HH:MM
    duration_min INTEGER DEFAULT 30,
    notes TEXT,
    FOREIGN KEY(client_id) REFERENCES clients(id) ON DELETE SET NULL
);
CREATE TABLE IF NOT EXISTS receipts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    number TEXT UNIQUE,
    client_id INTEGER,
    date TEXT NOT NULL,
    total_cents INTEGER NOT NULL DEFAULT 0,
    pdf_path TEXT,
    FOREIGN KEY(client_id) REFERENCES clients(id) ON DELETE SET NULL
);
CREATE TABLE IF NOT EXISTS receipt_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    receipt_id INTEGER,
    manipulation_id INTEGER,
    qty INTEGER DEFAULT 1,
    price_cents INTEGER NOT NULL,
    FOREIGN KEY(receipt_id) REFERENCES receipts(id) ON DELETE CASCADE,
    FOREIGN KEY(manipulation_id) REFERENCES manipulations(id)
);
"""

# ---- Helpers ----------------------------------------------------------------

def money_to_cents(x: float) -> int:
    return int(round(float(x) * 100))

def cents_to_money(c: int) -> str:
    return f"{c/100:.2f}"

@dataclass
class Company:
    name: str
    admin: str
    base_lang: str = "nl"

# ---- Data Layer -------------------------------------------------------------
class Store:
    def __init__(self, path=DB_PATH):
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self._init_db()

    def _init_db(self):
        cur = self.conn.cursor()
        cur.executescript(SCHEMA_SQL)
        self.conn.commit()

    # Config
    def get_config(self, key, default=None):
        cur = self.conn.cursor()
        cur.execute("SELECT value FROM config WHERE key=?", (key,))
        row = cur.fetchone()
        return row[0] if row else default

    def set_config(self, key, value):
        cur = self.conn.cursor()
        cur.execute("INSERT INTO config(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))
        self.conn.commit()

    # Company
    def get_company(self) -> Company | None:
        name = self.get_config("company_name")
        admin = self.get_config("admin_name")
        lang = self.get_config("base_lang", "nl")
        if name and admin:
            return Company(name=name, admin=admin, base_lang=lang)
        return None

    # Clients
    def add_client(self, name, email, phone, notes, lang="nl"):
        cur = self.conn.cursor()
        cur.execute("INSERT INTO clients(name,email,phone,notes,lang) VALUES(?,?,?,?,?)", (name,email,phone,notes,lang))
        self.conn.commit()

    def list_clients(self):
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM clients ORDER BY name")
        return cur.fetchall()

    def update_client(self, cid, name, email, phone, notes, lang):
        cur = self.conn.cursor()
        cur.execute("UPDATE clients SET name=?, email=?, phone=?, notes=?, lang=? WHERE id=?", (name,email,phone,notes,lang,cid))
        self.conn.commit()

    def delete_client(self, cid):
        cur = self.conn.cursor()
        cur.execute("DELETE FROM clients WHERE id=?", (cid,))
        self.conn.commit()

    # Manipulations
    def add_manip(self, name, price_cents):
        cur = self.conn.cursor()
        cur.execute("INSERT INTO manipulations(name, price_cents) VALUES(?,?)", (name, price_cents))
        self.conn.commit()

    def list_manips(self):
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM manipulations ORDER BY name")
        return cur.fetchall()

    def update_manip(self, mid, name, price_cents):
        cur = self.conn.cursor()
        cur.execute("UPDATE manipulations SET name=?, price_cents=? WHERE id=?", (name, price_cents, mid))
        self.conn.commit()

    def delete_manip(self, mid):
        cur = self.conn.cursor()
        cur.execute("DELETE FROM manipulations WHERE id=?", (mid,))
        self.conn.commit()

    # Appointments
    def add_appointment(self, client_id, date, time, duration_min, notes):
        cur = self.conn.cursor()
        cur.execute("INSERT INTO appointments(client_id,date,time,duration_min,notes) VALUES(?,?,?,?,?)", (client_id,date,time,duration_min,notes))
        self.conn.commit()

    def list_appointments_in_range(self, start_date: dt.date, end_date: dt.date):
        cur = self.conn.cursor()
        cur.execute(
            """
            SELECT a.*, c.name as client_name FROM appointments a
            LEFT JOIN clients c ON c.id=a.client_id
            WHERE date>=? AND date<=? ORDER BY date,time
            """,
            (start_date.isoformat(), end_date.isoformat()),
        )
        return cur.fetchall()

    def delete_appointment(self, aid):
        cur = self.conn.cursor()
        cur.execute("DELETE FROM appointments WHERE id=?", (aid,))
        self.conn.commit()

    # Receipts
    def create_receipt(self, client_id, items: list[tuple[int,int,int]]):
        """items: list of (manipulation_id, qty, price_cents)"""
        cur = self.conn.cursor()
        today = dt.date.today().isoformat()
        cur.execute("SELECT COUNT(*) FROM receipts WHERE date=?", (today,))
        count = cur.fetchone()[0] + 1
        number = f"{today.replace('-','')}-{count:04d}"
        total = sum(qty * price for _mid, qty, price in items)
        cur.execute("INSERT INTO receipts(number,client_id,date,total_cents) VALUES(?,?,?,?)", (number, client_id, today, total))
        rid = cur.lastrowid
        for mid, qty, price in items:
            cur.execute("INSERT INTO receipt_items(receipt_id, manipulation_id, qty, price_cents) VALUES(?,?,?,?)", (rid, mid, qty, price))
        self.conn.commit()
        return rid, number, total

    def get_receipt(self, rid):
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM receipts WHERE id=?", (rid,))
        r = cur.fetchone()
        cur.execute(
            """
            SELECT ri.*, m.name FROM receipt_items ri
            LEFT JOIN manipulations m ON m.id=ri.manipulation_id
            WHERE ri.receipt_id=?
            """,
            (rid,),
        )
        items = cur.fetchall()
        return r, items

    def update_receipt_pdf(self, rid, path):
        cur = self.conn.cursor()
        cur.execute("UPDATE receipts SET pdf_path=? WHERE id=?", (path, rid))
        self.conn.commit()

    def list_receipts_in_range(self, start: dt.date, end: dt.date):
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM receipts WHERE date>=? AND date<=? ORDER BY date", (start.isoformat(), end.isoformat()))
        return cur.fetchall()

    def list_receipts_in_range_by_client(self, start: dt.date, end: dt.date, client_id: int | None):
        cur = self.conn.cursor()
        if client_id:
            cur.execute("SELECT * FROM receipts WHERE date>=? AND date<=? AND client_id=? ORDER BY date", (start.isoformat(), end.isoformat(), client_id))
        else:
            cur.execute("SELECT * FROM receipts WHERE date>=? AND date<=? ORDER BY date", (start.isoformat(), end.isoformat()))
        return cur.fetchall()

    def sum_total_in_range(self, start: dt.date, end: dt.date) -> int:
        cur = self.conn.cursor()
        cur.execute("SELECT COALESCE(SUM(total_cents),0) FROM receipts WHERE date>=? AND date<=?", (start.isoformat(), end.isoformat()))
        return cur.fetchone()[0] or 0

    def sum_by_manipulations_in_range(self, start: dt.date, end: dt.date, client_id: int | None):
        cur = self.conn.cursor()
        if client_id:
            cur.execute(
                """
                SELECT m.name, COALESCE(SUM(ri.qty*ri.price_cents),0) as cents
                FROM receipts r
                JOIN receipt_items ri ON ri.receipt_id=r.id
                JOIN manipulations m ON m.id=ri.manipulation_id
                WHERE r.date>=? AND r.date<=? AND r.client_id=?
                GROUP BY m.name
                ORDER BY m.name
                """,
                (start.isoformat(), end.isoformat(), client_id),
            )
        else:
            cur.execute(
                """
                SELECT m.name, COALESCE(SUM(ri.qty*ri.price_cents),0) as cents
                FROM receipts r
                JOIN receipt_items ri ON ri.receipt_id=r.id
                JOIN manipulations m ON m.id=ri.manipulation_id
                WHERE r.date>=? AND r.date<=?
                GROUP BY m.name
                ORDER BY m.name
                """,
                (start.isoformat(), end.isoformat()),
            )
        return [(row[0], row[1]) for row in cur.fetchall()]

# ---- UI ---------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.store = Store()
        self.company = self.store.get_company()
        self.lang = self.company.base_lang if self.company else "nl"
        self.title(self.tr("app_title"))
        self.geometry("1120x780")
        self.minsize(1000, 680)

        if not self.company:
            self.first_run_wizard()

        self._build_menu()
        self._build_tabs()
        self.refresh_totals()

    def tr(self, key):
        return T.get(self.lang, T["nl"]).get(key, key)

    def _build_menu(self):
        menubar = tk.Menu(self)
        # Language
        lang_menu = tk.Menu(menubar, tearoff=0)
        for code in SUPPORTED_LANGS:
            lang_menu.add_command(label=code, command=lambda c=code: self.set_lang(c))
        menubar.add_cascade(label=self.tr("language"), menu=lang_menu)
        # Settings
        settings = tk.Menu(menubar, tearoff=0)
        settings.add_command(label=self.tr("set_vat"), command=self.set_vat_dialog)
        menubar.add_cascade(label=self.tr("settings"), menu=settings)
        self.config(menu=menubar)

    def set_lang(self, code):
        if code in SUPPORTED_LANGS:
            self.lang = code
            self.store.set_config("base_lang", code)
            self.title(self.tr("app_title"))
            for tab in (self.tab_dashboard, self.tab_agenda, self.tab_clients, self.tab_prices, self.tab_receipts):
                if hasattr(tab, 'refresh_labels'):
                    tab.refresh_labels()
            self._build_menu()

    def set_vat_dialog(self):
        cur = self.store.get_config("vat_rate", "21")
        val = simpledialog.askstring(self.tr("set_vat"), f"{self.tr('vat')} %:", initialvalue=str(cur))
        try:
            if val is not None:
                nv = float(val)
                if nv < 0 or nv > 100:
                    raise ValueError
                self.store.set_config("vat_rate", val)
        except Exception:
            messagebox.showerror(self.tr("set_vat"), "Ongeldige waarde")

    def first_run_wizard(self):
        messagebox.showinfo(self.tr("first_run_title"), self.tr("enter_company"))
        cname = simpledialog.askstring(self.tr("company"), self.tr("enter_company")) or "Mijn Pedicure"
        aname = simpledialog.askstring(self.tr("admin"), self.tr("enter_admin")) or "Administrator"
        blang = simpledialog.askstring(self.tr("language"), self.tr("choose_lang")) or "nl"
        if blang not in SUPPORTED_LANGS:
            blang = "nl"
        self.store.set_config("company_name", cname)
        self.store.set_config("admin_name", aname)
        self.store.set_config("base_lang", blang)
        self.store.set_config("vat_rate", "21")
        self.company = self.store.get_company()
        self.lang = blang

        messagebox.showinfo(self.tr("first_run_title"), self.tr("enter_manip_list"))
        defaults = [
            ("Basis pedicure", 35.0),
            ("Nagelknippen", 15.0),
            ("Eelt verwijderen", 20.0),
        ]
        for name, price in defaults:
            self.store.add_manip(name, money_to_cents(price))

    def _build_tabs(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill=tk.BOTH, expand=True)

        self.tab_dashboard = DashboardTab(self)
        self.tab_agenda = AgendaTab(self)
        self.tab_clients = ClientsTab(self)
        self.tab_prices = PricesTab(self)
        self.tab_receipts = ReceiptsTab(self)

        self.nb.add(self.tab_dashboard, text=self.tr("dashboard"))
        self.nb.add(self.tab_agenda, text=self.tr("agenda"))
        self.nb.add(self.tab_clients, text=self.tr("clients"))
        self.nb.add(self.tab_prices, text=self.tr("pricelist"))
        self.nb.add(self.tab_receipts, text=self.tr("receipts"))

    def refresh_totals(self):
        self.tab_dashboard.update_totals()

# ---- Dashboard ---------------------------------------------------------------
class DashboardTab(ttk.Frame):
    def __init__(self, app: App):
        super().__init__(app)
        self.app = app
        self.lbl_today = ttk.Label(self, font=("Segoe UI", 14))
        self.lbl_week = ttk.Label(self, font=("Segoe UI", 14))
        self.lbl_month = ttk.Label(self, font=("Segoe UI", 14))
        self.lbl_year = ttk.Label(self, font=("Segoe UI", 14))
        for w in (self.lbl_today, self.lbl_week, self.lbl_month, self.lbl_year):
            w.pack(anchor="w", padx=16, pady=8)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(anchor="w", padx=16, pady=8)
        self.btn_day = ttk.Button(btn_frame, text=self.app.tr("print_day"), command=self.print_day)
        self.btn_week = ttk.Button(btn_frame, text=self.app.tr("print_week"), command=self.print_week)
        self.btn_month = ttk.Button(btn_frame, text=self.app.tr("print_month"), command=self.print_month)
        self.btn_year = ttk.Button(btn_frame, text=self.app.tr("print_year"), command=self.print_year)
        for b in (self.btn_day, self.btn_week, self.btn_month, self.btn_year):
            b.pack(side=tk.LEFT, padx=6)

        self.refresh_labels()

    def refresh_labels(self):
        self.btn_day.config(text=self.app.tr("print_day"))
        self.btn_week.config(text=self.app.tr("print_week"))
        self.btn_month.config(text=self.app.tr("print_month"))
        self.btn_year.config(text=self.app.tr("print_year"))
        self.update_totals()

    def update_totals(self):
        store = self.app.store
        today = dt.date.today()
        start_week = today - dt.timedelta(days=today.weekday())
        end_week = start_week + dt.timedelta(days=6)
        start_month = today.replace(day=1)
        next_month = (start_month.replace(day=28) + dt.timedelta(days=4)).replace(day=1)
        end_month = next_month - dt.timedelta(days=1)
        start_year = today.replace(month=1, day=1)

        t_day = store.sum_total_in_range(today, today)
        t_week = store.sum_total_in_range(start_week, end_week)
        t_month = store.sum_total_in_range(start_month, end_month)
        t_year = store.sum_total_in_range(start_year, today)

        self.lbl_today.config(text=f"{self.app.tr('totals_today')}: € {cents_to_money(t_day)}")
        self.lbl_week.config(text=f"{self.app.tr('totals_week')}: € {cents_to_money(t_week)}")
        self.lbl_month.config(text=f"{self.app.tr('totals_month')}: € {cents_to_money(t_month)}")
        self.lbl_year.config(text=f"{self.app.tr('totals_year')}: € {cents_to_money(t_year)}")

    def _print_range(self, start: dt.date, end: dt.date, title: str):
        receipts = self.app.store.list_receipts_in_range(start, end)
        if not receipts:
            messagebox.showinfo(title, "Geen reçus in deze periode.")
            return
        if pdfcanvas is None:
            messagebox.showerror(title, self.app.tr("no_pdf"))
            return
        comp_slug = (self.app.company.name or "firma").lower().replace(" ", "_")
        fname = PDF_DIR / f"{comp_slug}_summary_{title.replace(' ', '_')}_{start}_{end}.pdf"
        c = pdfcanvas.Canvas(str(fname), pagesize=A4)
        width, height = A4
        c.setFont("Helvetica-Bold", 14)
        c.drawString(25*mm, height-25*mm, f"{self.app.company.name} – {title}")
        c.setFont("Helvetica", 10)
        y = height-35*mm
        total = 0
        for r in receipts:
            c.drawString(25*mm, y, f"{r['date']}  #{r['number']}  € {cents_to_money(r['total_cents'])}")
            y -= 6*mm
            total += r['total_cents']
            if y < 25*mm:
                c.showPage(); y = height-25*mm
        c.setFont("Helvetica-Bold", 12)
        c.drawString(25*mm, 20*mm, f"Totaal: € {cents_to_money(total)}")
        c.save()
        messagebox.showinfo(title, f"PDF opgeslagen: {fname}")

    def print_day(self):
        d = dt.date.today()
        self._print_range(d, d, self.app.tr("print_day"))

    def print_week(self):
        today = dt.date.today()
        start = today - dt.timedelta(days=today.weekday())
        end = start + dt.timedelta(days=6)
        self._print_range(start, end, self.app.tr("print_week"))

    def print_month(self):
        today = dt.date.today()
        start = today.replace(day=1)
        next_month = (start.replace(day=28) + dt.timedelta(days=4)).replace(day=1)
        end = next_month - dt.timedelta(days=1)
        self._print_range(start, end, self.app.tr("print_month"))

    def print_year(self):
        today = dt.date.today()
        start = today.replace(month=1, day=1)
        self._print_range(start, today, self.app.tr("print_year"))

# ---- Agenda -----------------------------------------------------------------
class AgendaTab(ttk.Frame):
    def __init__(self, app: App):
        super().__init__(app)
        self.app = app

        if Calendar is None:
            ttk.Label(self, text=self.app.tr("no_calendar"), foreground="red").pack(padx=16, pady=16)
            return

        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=6)
        ttk.Button(top, text=self.app.tr("new_appointment"), command=self.new_appointment).pack(side=tk.LEFT)

        body = ttk.Frame(self)
        body.pack(fill=tk.BOTH, expand=True)
        self.cal = Calendar(body, selectmode='day')
        self.cal.pack(side=tk.LEFT, fill=tk.Y, padx=8)

        right = ttk.Frame(body)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(right, columns=("date","time","client","notes"), show="headings")
        for col, w in (("date",140),("time",80),("client",220),("notes",400)):
            self.tree.heading(col, text=col.capitalize())
            self.tree.column(col, width=w)
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.cal.bind("<<CalendarSelected>>", lambda e: self.refresh_list())
        self.tree.bind("<Delete>", self.delete_selected)
        self.refresh_labels()
        self.refresh_list()

    def refresh_labels(self):
        self.tree.heading("date", text=self.app.tr("date"))
        self.tree.heading("time", text=self.app.tr("time"))
        self.tree.heading("client", text=self.app.tr("name"))
        self.tree.heading("notes", text=self.app.tr("notes"))

    def refresh_list(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        sel = self.cal.selection_get() if hasattr(self.cal,"selection_get") else dt.date.today()
        day = sel if isinstance(sel, dt.date) else dt.date.today()
        rows = self.app.store.list_appointments_in_range(day, day)
        for r in rows:
            self.tree.insert('', 'end', iid=r['id'], values=(r['date'], r['time'], r['client_name'] or "", r['notes'] or ""))

    def new_appointment(self):
        dlg = AppointmentDialog(self.app, self)
        self.wait_window(dlg)
        self.refresh_list()

    def delete_selected(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        if not messagebox.askyesno(self.app.tr("confirm"), self.app.tr("are_you_sure")):
            return
        for iid in sel:
            self.app.store.delete_appointment(int(iid))
        self.refresh_list()

class AppointmentDialog(tk.Toplevel):
    def __init__(self, app: App, parent):
        super().__init__(parent)
        self.app = app
        self.title(app.tr("new_appointment"))
        self.grab_set()

        frm = ttk.Frame(self)
        frm.pack(padx=12, pady=12)

        ttk.Label(frm, text=app.tr("date")).grid(row=0, column=0, sticky="e", padx=6, pady=4)
        if DateEntry:
            self.e_date = DateEntry(frm, date_pattern='yyyy-mm-dd')
        else:
            self.e_date = ttk.Entry(frm)
            self.e_date.insert(0, dt.date.today().isoformat())
        self.e_date.grid(row=0, column=1, sticky="w")

        ttk.Label(frm, text=app.tr("time")).grid(row=1, column=0, sticky="e", padx=6, pady=4)
        self.e_time = ttk.Entry(frm)
        self.e_time.grid(row=1, column=1, sticky="w")
        self.e_time.insert(0, "09:00")

        ttk.Label(frm, text=app.tr("duration_min")).grid(row=2, column=0, sticky="e", padx=6, pady=4)
        self.e_dur = ttk.Entry(frm)
        self.e_dur.grid(row=2, column=1, sticky="w")
        self.e_dur.insert(0, "30")

        ttk.Label(frm, text=app.tr("select_client")).grid(row=3, column=0, sticky="e", padx=6, pady=4)
        self.cb_client = ttk.Combobox(frm, values=[f"{c['id']}: {c['name']}" for c in app.store.list_clients()])
        self.cb_client.grid(row=3, column=1, sticky="w")

        ttk.Label(frm, text=app.tr("notes")).grid(row=4, column=0, sticky="e", padx=6, pady=4)
        self.e_notes = ttk.Entry(frm, width=40)
        self.e_notes.grid(row=4, column=1, sticky="w")

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=2, pady=10)
        ttk.Button(btns, text=app.tr("save"), command=self.save).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text=app.tr("cancel"), command=self.destroy).pack(side=tk.LEFT, padx=6)

    def save(self):
        date = self.e_date.get().strip()
        time = self.e_time.get().strip()
        dur = int(self.e_dur.get().strip() or 30)
        notes = self.e_notes.get().strip()
        cid = None
        if self.cb_client.get():
            try:
                cid = int(self.cb_client.get().split(":",1)[0])
            except:
                cid = None
        self.app.store.add_appointment(cid, date, time, dur, notes)
        self.destroy()

# ---- Clients ----------------------------------------------------------------
class ClientsTab(ttk.Frame):
    def __init__(self, app: App):
        super().__init__(app)
        self.app = app
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=6)
        ttk.Button(top, text=self.app.tr("add"), command=self.add).pack(side=tk.LEFT)
        ttk.Button(top, text=self.app.tr("edit"), command=self.edit).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text=self.app.tr("delete"), command=self.delete).pack(side=tk.LEFT, padx=4)

        self.tree = ttk.Treeview(self, columns=("name","email","phone","lang","notes"), show="headings")
        for col, w in (("name",200),("email",200),("phone",120),("lang",60),("notes",360)):
            self.tree.heading(col, text=col.capitalize())
            self.tree.column(col, width=w)
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.refresh_labels()
        self.refresh()

    def refresh_labels(self):
        self.tree.heading("name", text=self.app.tr("name"))
        self.tree.heading("email", text=self.app.tr("email"))
        self.tree.heading("phone", text=self.app.tr("phone"))
        self.tree.heading("lang", text=self.app.tr("client_lang"))
        self.tree.heading("notes", text=self.app.tr("notes"))

    def refresh(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for c in self.app.store.list_clients():
            self.tree.insert('', 'end', iid=c['id'], values=(c['name'], c['email'] or "", c['phone'] or "", c['lang'], c['notes'] or ""))

    def add(self):
        dlg = ClientDialog(self.app, self)
        self.wait_window(dlg)
        self.refresh()

    def edit(self):
        sel = self.tree.selection()
        if not sel:
            return
        cid = int(sel[0])
        row = [r for r in self.app.store.list_clients() if r['id']==cid][0]
        dlg = ClientDialog(self.app, self, existing=row)
        self.wait_window(dlg)
        self.refresh()

    def delete(self):
        sel = self.tree.selection()
        if not sel:
            return
        if not messagebox.askyesno(self.app.tr("confirm"), self.app.tr("are_you_sure")):
            return
        for iid in sel:
            self.app.store.delete_client(int(iid))
        self.refresh()

class ClientDialog(tk.Toplevel):
    def __init__(self, app: App, parent, existing=None):
        super().__init__(parent)
        self.app = app
        self.existing = existing
        self.title(app.tr("clients"))
        self.grab_set()

        frm = ttk.Frame(self)
        frm.pack(padx=12, pady=12)
        ttk.Label(frm, text=app.tr("name")).grid(row=0, column=0, sticky="e", padx=6, pady=4)
        ttk.Label(frm, text=app.tr("email")).grid(row=1, column=0, sticky="e", padx=6, pady=4)
        ttk.Label(frm, text=app.tr("phone")).grid(row=2, column=0, sticky="e", padx=6, pady=4)
        ttk.Label(frm, text=app.tr("client_lang")).grid(row=3, column=0, sticky="e", padx=6, pady=4)
        ttk.Label(frm, text=app.tr("notes")).grid(row=4, column=0, sticky="e", padx=6, pady=4)

        self.e_name = ttk.Entry(frm, width=40); self.e_name.grid(row=0, column=1, sticky="w")
        self.e_email = ttk.Entry(frm, width=40); self.e_email.grid(row=1, column=1, sticky="w")
        self.e_phone = ttk.Entry(frm, width=40); self.e_phone.grid(row=2, column=1, sticky="w")
        self.cb_lang = ttk.Combobox(frm, values=SUPPORTED_LANGS, width=8); self.cb_lang.grid(row=3, column=1, sticky="w")
        self.e_notes = ttk.Entry(frm, width=40); self.e_notes.grid(row=4, column=1, sticky="w")

        btns = ttk.Frame(frm); btns.grid(row=5, column=0, columnspan=2, pady=10)
        ttk.Button(btns, text=app.tr("save"), command=self.save).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text=app.tr("cancel"), command=self.destroy).pack(side=tk.LEFT, padx=6)

        if existing:
            self.e_name.insert(0, existing['name'])
            if existing['email']: self.e_email.insert(0, existing['email'])
            if existing['phone']: self.e_phone.insert(0, existing['phone'])
            self.cb_lang.set(existing['lang'] or 'nl')
            if existing['notes']: self.e_notes.insert(0, existing['notes'])
        else:
            self.cb_lang.set(self.app.lang)

    def save(self):
        name = self.e_name.get().strip()
        email = self.e_email.get().strip()
        phone = self.e_phone.get().strip()
        lang = self.cb_lang.get() or 'nl'
        notes = self.e_notes.get().strip()
        if not name:
            return
        if self.existing:
            self.app.store.update_client(self.existing['id'], name, email, phone, notes, lang)
        else:
            self.app.store.add_client(name, email, phone, notes, lang)
        self.destroy()

# ---- Prices -----------------------------------------------------------------
class PricesTab(ttk.Frame):
    def __init__(self, app: App):
        super().__init__(app)
        self.app = app
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=6)
        ttk.Button(top, text=self.app.tr("add"), command=self.add).pack(side=tk.LEFT)
        ttk.Button(top, text=self.app.tr("edit"), command=self.edit).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text=self.app.tr("delete"), command=self.delete).pack(side=tk.LEFT, padx=4)

        self.tree = ttk.Treeview(self, columns=("name","price"), show="headings")
        self.tree.heading("name", text=self.app.tr("manipulation"))
        self.tree.heading("price", text=self.app.tr("price"))
        self.tree.column("name", width=300)
        self.tree.column("price", width=100)
        self.tree.pack(fill=tk.BOTH, expand=True)

        btns = ttk.Frame(self)
        btns.pack(fill=tk.X, padx=8, pady=8)
        ttk.Button(btns, text=self.app.tr("change_manip"), command=self.edit).pack(side=tk.LEFT)
        ttk.Button(btns, text=self.app.tr("change_price"), command=self.edit).pack(side=tk.LEFT, padx=6)

        self.refresh_labels()
        self.refresh()

    def refresh_labels(self):
        self.tree.heading("name", text=self.app.tr("manipulation"))
        self.tree.heading("price", text=self.app.tr("price"))

    def refresh(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for m in self.app.store.list_manips():
            self.tree.insert('', 'end', iid=m['id'], values=(m['name'], f"€ {cents_to_money(m['price_cents'])}"))

    def add(self):
        dlg = ManipDialog(self.app, self)
        self.wait_window(dlg)
        self.refresh()

    def edit(self):
        sel = self.tree.selection()
        if not sel:
            return
        mid = int(sel[0])
        row = [r for r in self.app.store.list_manips() if r['id']==mid][0]
        dlg = ManipDialog(self.app, self, existing=row)
        self.wait_window(dlg)
        self.refresh()

    def delete(self):
        sel = self.tree.selection()
        if not sel:
            return
        if not messagebox.askyesno(self.app.tr("confirm"), self.app.tr("are_you_sure")):
            return
        for iid in sel:
            self.app.store.delete_manip(int(iid))
        self.refresh()

class ManipDialog(tk.Toplevel):
    def __init__(self, app: App, parent, existing=None):
        super().__init__(parent)
        self.app = app
        self.existing = existing
        self.title(app.tr("manipulation"))
        self.grab_set()

        frm = ttk.Frame(self)
        frm.pack(padx=12, pady=12)
        ttk.Label(frm, text=app.tr("manipulation")).grid(row=0, column=0, sticky="e", padx=6, pady=4)
        ttk.Label(frm, text=app.tr("price")).grid(row=1, column=0, sticky="e", padx=6, pady=4)
        self.e_name = ttk.Entry(frm, width=40); self.e_name.grid(row=0, column=1, sticky="w")
        self.e_price = ttk.Entry(frm, width=20); self.e_price.grid(row=1, column=1, sticky="w")

        btns = ttk.Frame(frm); btns.grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(btns, text=app.tr("save"), command=self.save).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text=app.tr("cancel"), command=self.destroy).pack(side=tk.LEFT, padx=6)

        if existing:
            self.e_name.insert(0, existing['name'])
            self.e_price.insert(0, str(existing['price_cents']/100))

    def save(self):
        name = self.e_name.get().strip()
        price_eur = float(self.e_price.get().strip() or 0)
        if not name:
            return
        if self.existing:
            self.app.store.update_manip(self.existing['id'], name, money_to_cents(price_eur))
        else:
            self.app.store.add_manip(name, money_to_cents(price_eur))
        self.destroy()

# ---- Receipts ----------------------------------------------------------------
class ReceiptsTab(ttk.Frame):
    def __init__(self, app: App):
        super().__init__(app)
        self.app = app
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=6)
        ttk.Button(top, text=self.app.tr("print_receipt"), command=self.new_receipt).pack(side=tk.LEFT)
        ttk.Button(top, text=self.app.tr("email_receipt"), command=self.email_selected).pack(side=tk.LEFT, padx=6)
        ttk.Button(top, text=self.app.tr("tax_doc"), command=self.print_tax_doc).pack(side=tk.LEFT, padx=12)

        # Client filter + periodeknoppen
        ttk.Label(top, text=self.app.tr("filter_client")).pack(side=tk.LEFT, padx=(16,4))
        self.clients = [{"id": None, "name": self.app.tr("all_clients")}]
        self.clients += list(self.app.store.list_clients())
        self.cb_filter_client = ttk.Combobox(top, values=[f"{c['id']}: {c['name']}" if c['id'] else c['name'] for c in self.clients], width=30)
        self.cb_filter_client.current(0)
        self.cb_filter_client.pack(side=tk.LEFT)
        self.cb_filter_client.bind("<<ComboboxSelected>>", lambda e: self.refresh())

        # Snelknoppen vaste perioden
        ttk.Button(top, text=self.app.tr("print_day"), command=lambda: self.print_period("day")).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text=self.app.tr("print_week"), command=lambda: self.print_period("week")).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text=self.app.tr("print_month"), command=lambda: self.print_period("month")).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text=self.app.tr("print_year"), command=lambda: self.print_period("year")).pack(side=tk.LEFT, padx=4)

        # Custom periode + export knoppen
        custom = ttk.Frame(self)
        custom.pack(fill=tk.X, padx=8, pady=6)
        ttk.Label(custom, text=self.app.tr("start_date")).pack(side=tk.LEFT)
        if DateEntry:
            self.e_from = DateEntry(custom, width=12, date_pattern='yyyy-mm-dd')
            self.e_from.set_date(dt.date.today().replace(day=1))
        else:
            self.e_from = ttk.Entry(custom, width=12)
            self.e_from.insert(0, dt.date.today().replace(day=1).isoformat())
        self.e_from.pack(side=tk.LEFT, padx=(4,12))

        ttk.Label(custom, text=self.app.tr("end_date")).pack(side=tk.LEFT)
        if DateEntry:
            self.e_to = DateEntry(custom, width=12, date_pattern='yyyy-mm-dd')
            self.e_to.set_date(dt.date.today())
        else:
            self.e_to = ttk.Entry(custom, width=12)
            self.e_to.insert(0, dt.date.today().isoformat())
        self.e_to.pack(side=tk.LEFT, padx=(4,12))

        ttk.Button(custom, text=self.app.tr("print_period"), command=self.print_custom).pack(side=tk.LEFT, padx=4)
        ttk.Button(custom, text=self.app.tr("export_csv"), command=self.export_csv).pack(side=tk.LEFT, padx=4)
        ttk.Button(custom, text=self.app.tr("export_excel"), command=self.export_excel).pack(side=tk.LEFT, padx=4)

        self.tree = ttk.Treeview(self, columns=("number","date","client","total","pdf"), show="headings")
        for col, w in (("number",160),("date",120),("client",240),("total",100),("pdf",380)):
            self.tree.heading(col, text=col.capitalize()); self.tree.column(col, width=w)
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.refresh_labels()
        self.refresh()

    def refresh_labels(self):
        self.tree.heading("number", text="#")
        self.tree.heading("date", text=self.app.tr("date"))
        self.tree.heading("client", text=self.app.tr("name"))
        self.tree.heading("total", text=self.app.tr("total"))

    def _selected_client_id(self):
        val = self.cb_filter_client.get()
        if not val or val == self.app.tr("all_clients"):
            return None
        try:
            return int(val.split(":",1)[0])
        except:
            return None

    def refresh(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        clients = {c['id']: c for c in self.app.store.list_clients()}
        rows = self.app.store.list_receipts_in_range_by_client(dt.date(1970,1,1), dt.date.today(), self._selected_client_id())
        for r in rows:
            cname = clients.get(r['client_id'], {}).get('name', '')
            self.tree.insert('', 'end', iid=r['id'], values=(r['number'], r['date'], cname, f"€ {cents_to_money(r['total_cents'])}", r['pdf_path'] or ""))

    def _period_dates(self, period: str):
        today = dt.date.today()
        if period == "day":
            return today, today
        if period == "week":
            start = today - dt.timedelta(days=today.weekday())
            end = start + dt.timedelta(days=6)
            return start, end
        if period == "month":
            start = today.replace(day=1)
            next_month = (start.replace(day=28) + dt.timedelta(days=4)).replace(day=1)
            end = next_month - dt.timedelta(days=1)
            return start, end
        if period == "year":
            start = today.replace(month=1, day=1)
            return start, today
        return today, today

    def _parse_date_or(self, s: str, default: dt.date) -> dt.date:
        try:
            return dt.date.fromisoformat(s.strip()) if isinstance(s, str) else s
        except Exception:
            return default

    def _print_range_receipts(self, start: dt.date, end: dt.date, client_id: int | None):
        receipts = self.app.store.list_receipts_in_range_by_client(start, end, client_id)
        if not receipts:
            messagebox.showinfo(self.app.tr("print_period"), "Geen reçus in deze periode.")
            return
        if pdfcanvas is None:
            messagebox.showerror(self.app.tr("print_period"), self.app.tr("no_pdf"))
            return
        # Bestandsnaam met firmanaam
        comp_slug = (self.app.company.name or "firma").lower().replace(" ", "_")
        cname = None
        if client_id:
            for c in self.app.store.list_clients():
                if c['id'] == client_id:
                    cname = c['name']
                    break
        suffix = f"{start}_{end}" + (f"_{cname}" if cname else "_ALL")
        fname = PDF_DIR / f"{comp_slug}_receipts_{suffix}.pdf"
        c = pdfcanvas.Canvas(str(fname), pagesize=A4)
        width, height = A4
        c.setFont("Helvetica-Bold", 14)
        title = f"{self.app.company.name} – {self.app.tr('receipts')} {start} → {end}" + (f" – {cname}" if cname else "")
        c.drawString(25*mm, height-25*mm, title)
        c.setFont("Helvetica", 10)
        y = height-35*mm
        total = 0
        for r in receipts:
            line = f"{r['date']}  #{r['number']}  € {cents_to_money(r['total_cents'])}"
            c.drawString(25*mm, y, line)
            y -= 6*mm
            total += r['total_cents']
            if y < 70*mm:
                c.showPage(); y = height-25*mm
        # Subtotalen per manipulatie
        sums = self.app.store.sum_by_manipulations_in_range(start, end, client_id)
        c.setFont("Helvetica-Bold", 11)
        c.drawString(25*mm, y-6*mm, self.app.tr("subtotal_by_manip"))
        y -= 12*mm
        c.setFont("Helvetica", 10)
        for name, cents in sums:
            c.drawString(25*mm, y, f"• {name}")
            c.drawRightString(170*mm, y, f"€ {cents_to_money(cents)}")
            y -= 6*mm
            if y < 40*mm:
                c.showPage(); y = height-40*mm
        # BTW uitsplitsing
        vat_rate = float(self.app.store.get_config("vat_rate", "21") or 21)
        vat_amount = int(round(total * (vat_rate / (100 + vat_rate)))) if vat_rate > 0 else 0
        net_amount = total - vat_amount
        c.setFont("Helvetica-Bold", 12)
        c.drawString(25*mm, 26*mm, f"Netto: € {cents_to_money(net_amount)}")
        c.drawString(25*mm, 20*mm, f"{self.app.tr('vat')} {vat_rate:.2f}%: € {cents_to_money(vat_amount)}")
        c.drawString(25*mm, 14*mm, f"Totaal: € {cents_to_money(total)}")
        c.save()
        messagebox.showinfo(self.app.tr("print_period"), f"PDF opgeslagen: {fname}")

    def print_period(self, period: str):
        start, end = self._period_dates(period)
        self._print_range_receipts(start, end, self._selected_client_id())

    def print_custom(self):
        today = dt.date.today()
        start = self._parse_date_or(self.e_from.get(), today)
        end = self._parse_date_or(self.e_to.get(), today)
        self._print_range_receipts(start, end, self._selected_client_id())

    def _rows_for_export(self, start: dt.date, end: dt.date):
        rows = self.app.store.list_receipts_in_range_by_client(start, end, self._selected_client_id())
        clients = {c['id']: c for c in self.app.store.list_clients()}
        data = []
        for r in rows:
            cname = clients.get(r['client_id'], {}).get('name', '')
            data.append({'number': r['number'], 'date': r['date'], 'client': cname, 'total_eur': cents_to_money(r['total_cents']), 'pdf_path': r['pdf_path'] or ''})
        return data

    def export_csv(self):
        import csv
        today = dt.date.today()
        start = self._parse_date_or(self.e_from.get(), today.replace(day=1))
        end = self._parse_date_or(self.e_to.get(), today)
        data = self._rows_for_export(start, end)
        if not data:
            messagebox.showinfo("CSV", "Geen data voor export.")
            return
        comp_slug = (self.app.company.name or "firma").lower().replace(" ", "_")
        fname = PDF_DIR / f"{comp_slug}_export_{start}_{end}.csv"
        with open(fname, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=['number','date','client','total_eur','pdf_path'])
            writer.writeheader(); writer.writerows(data)
        messagebox.showinfo("CSV", f"CSV opgeslagen: {fname}")

    def export_excel(self):
        try:
            import xlsxwriter  # type: ignore
        except Exception:
            messagebox.showerror("Excel", "Pakket 'xlsxwriter' ontbreekt. Installeer met: pip install xlsxwriter")
            return
        today = dt.date.today()
        start = self._parse_date_or(self.e_from.get(), today.replace(day=1))
        end = self._parse_date_or(self.e_to.get(), today)
        data = self._rows_for_export(start, end)
        if not data:
            messagebox.showinfo("Excel", "Geen data voor export.")
            return
        comp_slug = (self.app.company.name or "firma").lower().replace(" ", "_")
        fname = PDF_DIR / f"{comp_slug}_export_{start}_{end}.xlsx"
        wb = xlsxwriter.Workbook(str(fname)); ws = wb.add_worksheet('Receipts')
        headers = ['number','date','client','total_eur','pdf_path']
        for col, h in enumerate(headers): ws.write(0, col, h)
        for r, row in enumerate(data, start=1):
            ws.write(r, 0, row['number']); ws.write(r, 1, row['date']); ws.write(r, 2, row['client']); ws.write(r, 3, row['total_eur']); ws.write(r, 4, row['pdf_path'])
        ws.autofilter(0, 0, len(data), len(headers)-1); ws.set_column(0, len(headers)-1, 22)
        wb.close(); messagebox.showinfo("Excel", f"Excel opgeslagen: {fname}")

    def new_receipt(self):
        dlg = ReceiptDialog(self.app, self)
        self.wait_window(dlg)
        self.refresh(); self.app.refresh_totals()

    def email_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        rid = int(sel[0])
        r, items = self.app.store.get_receipt(rid)
        client = None
        for c in self.app.store.list_clients():
            if c['id'] == r['client_id']:
                client = c; break
        if not client or not client['email']:
            messagebox.showerror("Email", "Geen e-mail voor deze cliënt.")
            return
        send_receipt_email(self.app, r, items, client)

    def print_tax_doc(self):
        today = dt.date.today(); start = today.replace(month=1, day=1); end = today
        receipts = self.app.store.list_receipts_in_range_by_client(start, end, None)
        if not receipts:
            messagebox.showinfo("Belastingdocument", "Geen reçus dit jaar.")
            return
        if pdfcanvas is None:
            messagebox.showerror("Belastingdocument", self.app.tr("no_pdf"))
            return
        try:
            vat_rate = float(self.app.store.get_config("vat_rate", "21") or 21)
        except Exception:
            vat_rate = 21.0
        comp_slug = (self.app.company.name or "firma").lower().replace(" ", "_")
        fname = PDF_DIR / f"{comp_slug}_tax_declaration_{today.year}.pdf"
        c = pdfcanvas.Canvas(str(fname), pagesize=A4)
        width, height = A4
        c.setFont("Helvetica-Bold", 14)
        c.drawString(25*mm, height-25*mm, f"{self.app.company.name} – Jaaroverzicht {today.year}")
        c.setFont("Helvetica", 10)
        c.drawString(25*mm, height-32*mm, f"Administrator: {self.app.company.admin}")
        c.drawString(25*mm, height-38*mm, f"Prijzen inclusief btw ({vat_rate:.2f}%).")
        month_totals = {}; year_total = 0
        for r in receipts:
            d = dt.date.fromisoformat(r['date']); key = f"{d.year}-{d.month:02d}"
            month_totals[key] = month_totals.get(key, 0) + r['total_cents']
            year_total += r['total_cents']
        y = height-52*mm
        c.setFont("Helvetica-Bold", 11)
        c.drawString(25*mm, y, "Maand"); c.drawRightString(115*mm, y, "Netto (€)"); c.drawRightString(150*mm, y, f"Btw {vat_rate:.2f}% (€)"); c.drawRightString(170*mm, y, "Totaal (€)")
        y -= 6*mm; c.setLineWidth(0.5); c.line(25*mm, y, 170*mm, y); y -= 4*mm
        c.setFont("Helvetica", 10); year_vat = 0; year_net = 0
        for m in sorted(month_totals.keys()):
            gross = month_totals[m]
            vat_amt = int(round(gross * (vat_rate / (100.0 + vat_rate)))) if vat_rate > 0 else 0
            net_amt = gross - vat_amt
            year_vat += vat_amt; year_net += net_amt
            c.drawString(25*mm, y, m)
            c.drawRightString(115*mm, y, cents_to_money(net_amt))
            c.drawRightString(150*mm, y, cents_to_money(vat_amt))
            c.drawRightString(170*mm, y, cents_to_money(gross))
            y -= 6*mm
            if y < 30*mm:
                c.showPage(); y = height-25*mm
        if y < 40*mm:
            c.showPage(); y = height-25*mm
        c.setLineWidth(0.5); c.line(25*mm, y, 170*mm, y); y -= 8*mm
        c.setFont("Helvetica-Bold", 12); c.drawString(25*mm, y, "Jaar totalen"); y -= 8*mm
        c.setFont("Helvetica", 11)
        c.drawString(30*mm, y, "Netto"); c.drawRightString(80*mm, y, f"€ {cents_to_money(year_net)}")
        c.drawString(95*mm, y, f"Btw {vat_rate:.2f}%"); c.drawRightString(150*mm, y, f"€ {cents_to_money(year_vat)}")
        y -= 7*mm; c.setFont("Helvetica-Bold", 12)
        c.drawString(30*mm, y, "Totaal"); c.drawRightString(170*mm, y, f"€ {cents_to_money(year_total)}")
        c.save(); messagebox.showinfo("Belastingdocument", f"PDF opgeslagen: {fname}")

class ReceiptDialog(tk.Toplevel):
    def __init__(self, app: App, parent):
        super().__init__(parent)
        self.app = app
        self.title(app.tr("receipt"))
        self.grab_set()

        frm = ttk.Frame(self); frm.pack(padx=12, pady=12)
        ttk.Label(frm, text=app.tr("select_client")).grid(row=0, column=0, sticky="e", padx=6, pady=4)
        self.clients = app.store.list_clients()
        self.cb_client = ttk.Combobox(frm, values=[f"{c['id']}: {c['name']}" for c in self.clients], width=40)
        self.cb_client.grid(row=0, column=1, sticky="w")
        ttk.Label(frm, text=app.tr("select_manips")).grid(row=1, column=0, sticky="ne", padx=6, pady=4)
        self.manips = app.store.list_manips()
        self.lb_manips = tk.Listbox(frm, selectmode=tk.MULTIPLE, width=50, height=10)
        for m in self.manips:
            self.lb_manips.insert(tk.END, f"{m['id']} | {m['name']} | € {cents_to_money(m['price_cents'])}")
        self.lb_manips.grid(row=1, column=1, sticky="w")
        self.lbl_total = ttk.Label(frm, text=f"{app.tr('total')}: € 0.00", font=("Segoe UI", 12, "bold"))
        self.lbl_total.grid(row=2, column=1, sticky="w", pady=6)
        btns = ttk.Frame(frm); btns.grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(btns, text=app.tr("print_receipt"), command=self.save).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text=app.tr("cancel"), command=self.destroy).pack(side=tk.LEFT, padx=6)
        self.lb_manips.bind('<<ListboxSelect>>', self.update_total)

    def update_total(self, event=None):
        total = 0
        for idx in self.lb_manips.curselection():
            m = self.manips[idx]; total += m['price_cents']
        self.lbl_total.config(text=f"{self.app.tr('total')}: € {cents_to_money(total)}")

    def save(self):
        if not self.cb_client.get():
            return
        cid = int(self.cb_client.get().split(":",1)[0])
        items = []
        for idx in self.lb_manips.curselection():
            m = self.manips[idx]
            items.append((m['id'], 1, m['price_cents']))
        if not items:
            return
        rid, number, total = self.app.store.create_receipt(cid, items)
        path = generate_receipt_pdf(self.app, rid)
        if path:
            self.app.store.update_receipt_pdf(rid, str(path))
        messagebox.showinfo(self.app.tr("receipt"), f"Reçu #{number} – € {cents_to_money(total)}")
        self.destroy()

# ---- PDF & Email ------------------------------------------------------------

def generate_receipt_pdf(app: App, rid: int) -> Path | None:
    if pdfcanvas is None:
        messagebox.showerror(app.tr("receipt"), app.tr("no_pdf"))
        return None
    company = app.company
    r, items = app.store.get_receipt(rid)
    client = None
    for c in app.store.list_clients():
        if c['id'] == r['client_id']:
            client = c; break
    try:
        vat_rate = float(app.store.get_config("vat_rate", "21") or 21)
    except Exception:
        vat_rate = 21.0
    comp_slug = (company.name or "firma").lower().replace(" ", "_")
    fname = PDF_DIR / f"{comp_slug}_receipt_{r['number']}.pdf"
    c = pdfcanvas.Canvas(str(fname), pagesize=A4)
    width, height = A4
    c.setStrokeColor(colors.black)
    c.rect(20*mm, height-30*mm, width-40*mm, 20*mm, stroke=1, fill=0)
    c.setFont("Helvetica-Bold", 16); c.drawString(25*mm, height-18*mm, company.name)
    c.setFont("Helvetica", 10); c.drawString(25*mm, height-24*mm, f"{app.tr('admin')}: {company.admin}")
    c.setFont("Helvetica-Bold", 12); c.drawString(25*mm, height-40*mm, f"{app.tr('receipt')} #{r['number']}")
    c.setFont("Helvetica", 10); c.drawString(25*mm, height-46*mm, f"{app.tr('date')}: {r['date']}")
    if client:
        c.drawString(25*mm, height-52*mm, f"{app.tr('name')}: {client['name']}")
        if client['email']:
            c.drawString(25*mm, height-58*mm, f"{app.tr('email')}: {client['email']}")
    y = height-72*mm
    c.setFont("Helvetica-Bold", 10)
    c.drawString(25*mm, y, app.tr("manipulation"))
    c.drawRightString(170*mm, y, app.tr("price"))
    y -= 6*mm; c.setFont("Helvetica", 10)
    total = 0
    for it in items:
        name = it['name']; price = it['price_cents'] * it['qty']; total += price
        c.drawString(25*mm, y, f"{name}"); c.drawRightString(170*mm, y, f"€ {cents_to_money(price)}")
        y -= 6*mm
        if y < 40*mm:
            c.showPage(); y = height-30*mm
    vat_amount = int(round(total * (vat_rate / (100.0 + vat_rate)))) if vat_rate > 0 else 0
    net_amount = total - vat_amount
    if y < 30*mm:
        c.showPage(); y = height-30*mm
    c.setFont("Helvetica-Bold", 11)
    c.drawRightString(170*mm, y-2*mm, f"{app.tr('total')}: € {cents_to_money(total)}")
    y -= 10*mm; c.setFont("Helvetica", 10)
    c.drawRightString(170*mm, y, f"Netto: € {cents_to_money(net_amount)}"); y -= 6*mm
    c.drawRightString(170*mm, y, f"Btw {vat_rate:.2f}%: € {cents_to_money(vat_amount)}"); y -= 8*mm
    c.setLineWidth(0.5); c.line(120*mm, y, 170*mm, y); y -= 6*mm
    c.setFont("Helvetica-Bold", 11); c.drawRightString(170*mm, y, f"Totaal incl. btw: € {cents_to_money(total)}")
    y -= 10*mm; c.setFont("Helvetica", 8)
    c.drawString(25*mm, y, "Prijzen inclusief btw. Bewaar dit reçu voor uw administratie.")
    c.save(); return fname


def send_receipt_email(app: App, r, items, client_row):
    import smtplib
    from email.message import EmailMessage
    SMTP_HOST = os.getenv("SMTP_HOST", ""); SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
    SMTP_USER = os.getenv("SMTP_USER", ""); SMTP_PASS = os.getenv("SMTP_PASS", ""); SMTP_FROM = os.getenv("SMTP_FROM", SMTP_USER)
    if not all([SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_FROM]):
        messagebox.showerror("Email", "SMTP instellingen ontbreken (env: SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_FROM)"); return
    if not r['pdf_path'] or not Path(r['pdf_path']).exists():
        path = generate_receipt_pdf(app, r['id'])
        if path:
            app.store.update_receipt_pdf(r['id'], str(path))
    filepath = r['pdf_path']
    lang = client_row['lang'] or app.lang
    subj_map = {'nl': f"Reçu #{r['number']} – {app.company.name}", 'fr': f"Reçu #{r['number']} – {app.company.name}", 'en': f"Receipt #{r['number']} – {app.company.name}", 'ar': f"إيصال #{r['number']} – {app.company.name}"}
    body_map = {
        'nl': f"Beste {client_row['name']},\nIn de bijlage vindt u uw reçu. Dank u.\n{app.company.name}",
        'fr': f"Cher/Chère {client_row['name']},\nVeuillez trouver en pièce jointe votre reçu. Merci.\n{app.company.name}",
        'en': f"Dear {client_row['name']},\nPlease find your receipt attached. Thank you.\n{app.company.name}",
        'ar': f"{client_row['name']} العزيز/العزيزة،\nيرجى العثور على الإيصال بالمرفق. شكراً لك.\n{app.company.name}",
    }
    subj = subj_map.get(lang, subj_map['en']); body = body_map.get(lang, body_map['en'])
    msg = EmailMessage(); msg['Subject'] = subj; msg['From'] = SMTP_FROM; msg['To'] = client_row['email']; msg.set_content(body)
    if filepath and Path(filepath).exists():
        with open(filepath, 'rb') as f:
            data = f.read(); msg.add_attachment(data, maintype='application', subtype='pdf', filename=Path(filepath).name)
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
            s.starttls(); s.login(SMTP_USER, SMTP_PASS); s.send_message(msg)
        messagebox.showinfo("Email", "E-mail verzonden.")
    except Exception as e:
        messagebox.showerror("Email", f"Fout bij verzenden: {e}")

# ---- main -------------------------------------------------------------------
if __name__ == "__main__":
    app = App(); app.mainloop()
