# -*- coding: utf-8 -*-
import base64
import json
import smtplib
import ssl
import subprocess
import tempfile
import email.utils
from email.message import EmailMessage
import mimetypes
import shutil
import time
import re
from io import BytesIO
import threading
from datetime import datetime, date, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs, unquote
from zoneinfo import ZoneInfo

import solar_invoice as s

ROOT = Path(__file__).resolve().parent
OUTPUT = s.OUTPUT
HOST = "127.0.0.1"
PORT = 8774
APP = ROOT / "Solarabrechnung.html"
MAIL_CRED_TARGET = s.CRED_TARGET + ".SmtpMailboxPassword"
SETTINGS_FILE = ROOT / "abrechnungseinstellungen.json"


class UploadedFile:
    def __init__(self, data, name):
        self._buffer = BytesIO(data)
        self.name = name

    def __getattr__(self, name):
        return getattr(self._buffer, name)

def load_history():
    path = OUTPUT / "history.json"
    if not path.exists():
        return []
    data = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(data, dict):
        data = [data]
        s.history_index(data)
    if not isinstance(data, list):
        raise RuntimeError("Das Rechnungsregister hat ein ungültiges Format.")
    return data

def load_settings():
    default = {"grid_rp": s.CFG["grid_rp"], "solar_rp": s.CFG["solar_rp"], "reduction_mode": "percent", "reduction_value": 0}
    if SETTINGS_FILE.exists():
        try:
            default.update(json.loads(SETTINGS_FILE.read_text(encoding="utf-8")))
        except (OSError, json.JSONDecodeError):
            pass
    return default


def save_settings(data):
    current = load_settings()
    for key in ("grid_rp", "solar_rp", "reduction_value"):
        if key in data:
            current[key] = float(data[key])
    if data.get("reduction_mode") in ("percent", "fixed"):
        current["reduction_mode"] = data["reduction_mode"]
    if current["grid_rp"] < 0 or current["solar_rp"] < 0 or current["reduction_value"] < 0:
        raise RuntimeError("Tarife und Reduktion dürfen nicht negativ sein.")
    if current["reduction_mode"] == "percent" and current["reduction_value"] > 100:
        raise RuntimeError("Der Rabatt darf höchstens 100 % betragen.")
    SETTINGS_FILE.write_text(json.dumps(current, ensure_ascii=False, indent=2), encoding="utf-8")
    return current

def generate(files, key, settings=None):
    required = ("iban", "sender", "street", "house", "postal", "city")
    missing = [name for name in required if not str(s.CFG.get(name, "")).strip()]
    if missing:
        raise RuntimeError("Bitte zuerst im Abschnitt «Absender und Zahlung» die eigenen Angaben speichern.")
    if not isinstance(s.CFG.get("customers"), dict) or not s.CFG["customers"]:
        raise RuntimeError("Bitte im Abschnitt «Bezüger und Zähler» mindestens einen Bezüger mit Zählernummer erfassen und speichern.")
    if key:
        s.cred_write(key)
    settings = save_settings(settings or load_settings())
    grid_rp = float(settings.get("grid_rp", s.CFG["grid_rp"]))
    solar_rp = float(settings.get("solar_rp", s.CFG["solar_rp"]))
    reduction_mode = str(settings.get("reduction_mode", "percent"))
    reduction_value = float(settings.get("reduction_value", 0) or 0)
    if grid_rp < 0 or solar_rp < 0 or reduction_value < 0:
        raise RuntimeError("Tarife und Reduktion dürfen nicht negativ sein.")
    if reduction_mode not in ("percent", "fixed"):
        raise RuntimeError("Unbekannte Reduktionsart.")
    if reduction_mode == "percent" and reduction_value > 100:
        raise RuntimeError("Der prozentuale Rabatt darf höchstens 100 % betragen.")
    records = []
    for item in files:
        name = Path(item.get("name", "smartme.xlsx")).name
        upload = UploadedFile(base64.b64decode(item["data"]), name)
        records.append(s.read_export(upload))
        upload.close()
    if not records:
        raise RuntimeError("Bitte mindestens einen Smart-me-Export auswählen.")
    first = records[0]
    year, quarter = s.quarter(first["start"], first["end"])
    if any((x["start"], x["end"]) != (first["start"], first["end"]) for x in records):
        raise RuntimeError("Alle Smart-me-Dateien müssen denselben Quartalszeitraum enthalten.")
    if len({x["meter"] for x in records}) != len(records):
        raise RuntimeError("Ein Smart-me-Zähler wurde mehrfach ausgewählt.")

    token = s.api_token()
    smid = s.gateway(token)
    query = s.urllib.parse.urlencode({
        "accuracy": "low",
        "from": s.utc_iso(first["start"]),
        "to": s.utc_iso(first["end"], True),
    })
    solar = s.call(
        f'{s.API}/v1/statistics/gateways/{s.urllib.parse.quote(smid, safe="")}?{query}',
        token=token,
    )
    if not all(k in solar for k in ("production", "consumption", "selfConsumption", "selfConsumptionRate", "autarchyDegree")):
        raise RuntimeError("Solar-Manager-Kennzahlen fehlen oder die API-Antwort ist unvollständig.")

    month_names = ("Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember")
    monthly_production = []
    month_start = first["start"]
    while month_start <= first["end"]:
        next_month = date(month_start.year + (month_start.month == 12), month_start.month % 12 + 1, 1)
        month_end = min(next_month - timedelta(days=1), first["end"])
        month_solar = s.stats(token, smid, month_start, month_end)
        monthly_production.append({"month": month_names[month_start.month - 1], "production": float(month_solar["production"]) / 1000})
        month_start = next_month
    solar["monthlyProduction"] = monthly_production

    today = datetime.now(ZoneInfo("Europe/Zurich")).date()
    period = f"{year} Q{quarter}"
    folder = OUTPUT / f"{year}-Q{quarter}"
    folder.mkdir(parents=True, exist_ok=True)
    history_path = OUTPUT / "history.json"
    entries = load_history()
    seq = len(entries)
    generated = []
    smartme_total_kwh = sum(float(r["kwh"]) for r in records)
    solar_manager_consumption_kwh = float(solar["consumption"]) / 1000
    solar_self_consumption_kwh = float(solar["selfConsumption"]) / 1000
    if solar_manager_consumption_kwh <= 0:
        raise RuntimeError("Solar Manager meldet keinen Gesamtverbrauch für den Abrechnungszeitraum.")
    if solar_self_consumption_kwh > solar_manager_consumption_kwh + 0.01:
        raise RuntimeError("Der Solar-Manager-Eigenverbrauch ist höher als der Gesamtverbrauch. Bitte Messwerte und Zeitraum prüfen.")
    allocation_fraction = min(1.0, solar_self_consumption_kwh / solar_manager_consumption_kwh)
    for record in records:
        customer = s.CFG["customers"][record["meter"]]
        solar_kwh = record["kwh"] * allocation_fraction
        grid_kwh = record["kwh"] - solar_kwh
        grid_amount = round(grid_kwh * grid_rp / 100, 2)
        solar_amount = round(solar_kwh * solar_rp / 100, 2)
        subtotal = round(grid_amount + solar_amount, 2)
        if reduction_mode == "percent":
            reduction_amount = round(subtotal * reduction_value / 100, 2)
            reduction_description = f"Reduktion ({reduction_value:.2f} %)"
        else:
            reduction_amount = round(min(subtotal, reduction_value), 2)
            reduction_description = f"Pauschalabzug CHF {reduction_amount:.2f}"
        meter_rent = float(s.CFG.get("meter_rent_chf", 8.00))
        amount = round(subtotal - reduction_amount + meter_rent, 2)
        seq += 1
        number = f"{year}Q{quarter}-{seq:04d}"
        digits = "".join(c for c in number if c.isdigit())[-21:] or "1"
        check = 98 - (int(digits + "271500") % 97)
        reference = f"RF{check:02d}{digits}"
        invoice = {
            "no": number, "ref": reference, "date": today,
            "due": today + timedelta(days=s.CFG["due_days"]),
            "start": record["start"], "end": record["end"],
            "period": period, "customer": customer, "cons": record["kwh"],
            "solar": solar_kwh, "grid": grid_kwh, "solar_allocation_pct": allocation_fraction * 100,
            "grid_amt": grid_amount, "solar_amt": solar_amount,
            "amount": amount, "grid_rp": grid_rp, "solar_rp": solar_rp,
            "subtotal": subtotal, "reduction_amount": reduction_amount, "meter_rent": meter_rent,
            "reduction_description": reduction_description,
            "savings": max(0, record["kwh"] * grid_rp / 100 + meter_rent - amount),
        }
        path = folder / f"Rechnung-{record['meter']}-{year}-Q{quarter}-{seq:04d}.html"
        path.write_text(s.render(invoice, solar), encoding="utf-8")
        entry = {
            "invoice": number, "customer": customer["name"], "email": customer.get("email", ""), "meter": record["meter"], "solar_allocation_pct": allocation_fraction * 100,
            "period": period, "issued": today.isoformat(), "amount": amount,
            "path": path.relative_to(OUTPUT).as_posix(),
        }
        entries.append(entry)
        generated.append(entry)
    s.history_index(entries)
    solar_consumption_kwh = solar_manager_consumption_kwh
    smartme_total_kwh = sum(float(r["kwh"]) for r in records)
    variance_kwh = smartme_total_kwh - solar_consumption_kwh
    variance_pct = abs(variance_kwh) / solar_consumption_kwh * 100 if solar_consumption_kwh else None
    return {
        "generated": generated, "history": entries, "period": period, "solar": solar,
        "smartme_total_kwh": smartme_total_kwh, "solar_allocation_pct": allocation_fraction * 100,
        "solar_manager_consumption_kwh": solar_consumption_kwh,
        "consumption_variance_kwh": variance_kwh,
        "consumption_variance_pct": variance_pct,
    }


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        # Avoid logging request bodies or credential material.
        return

    def respond(self, status, obj, content_type="application/json; charset=utf-8"):
        body = obj if isinstance(obj, bytes) else json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            return self.serve_file(APP)
        if parsed.path == "/api/status":
            self.respond(200, {"key_saved": bool(s.cred_read()), "mail_saved": bool(s.cred_read(MAIL_CRED_TARGET))})
            return
        if parsed.path == "/api/settings":
            self.respond(200, load_settings())
            return
        if parsed.path == "/api/profile":
            fields = ("iban", "sender", "street", "house", "postal", "city")
            self.respond(200, {name: s.CFG.get(name, "") for name in fields})
            return
        if parsed.path == "/api/customers":
            self.respond(200, [{"meterId": meter_id, **customer} for meter_id, customer in s.CFG.get("customers", {}).items()])
            return
        if parsed.path == "/api/history":
            p = OUTPUT / "history.json"
            data = load_history()
            self.respond(200, data)
            return
        if parsed.path.startswith("/files/"):
            relative = unquote(parsed.path[len("/files/"):])
            target = (OUTPUT / relative).resolve()
            if not target.is_relative_to(OUTPUT.resolve()) or not target.is_file():
                self.respond(404, {"error": "Datei nicht gefunden"})
                return
            return self.serve_file(target)
        self.respond(404, {"error": "Nicht gefunden"})

    def serve_file(self, path):
        data = path.read_bytes()
        kind = mimetypes.guess_type(str(path))[0] or "application/octet-stream"
        if kind == "text/html":
            kind += "; charset=utf-8"
        self.respond(200, data, kind)

    def do_POST(self):
        parsed = urlparse(self.path)
        try:
            size = int(self.headers.get("Content-Length", "0"))
            if size > 80_000_000:
                raise RuntimeError("Die ausgewählten Dateien sind zusammen zu gross.")
            payload = json.loads(self.rfile.read(size).decode("utf-8"))
            if parsed.path == "/api/profile":
                self.respond(200, s.save_profile(payload))
                return
            if parsed.path == "/api/customers":
                self.respond(200, s.save_customers(payload.get("customers", [])))
                return
            if parsed.path == "/api/key":
                key = (payload.get("apiKey") or "").strip()
                if not key:
                    raise RuntimeError("Kein API-Schlüssel eingegeben.")
                s.cred_write(key)
                self.respond(200, {"saved": True})
                return
            if parsed.path == "/api/mail-settings":
                password = str(payload.get("password", "")).strip()
                if len(password) < 4:
                    raise RuntimeError("Bitte das Mailbox-Passwort eingeben.")
                s.cred_write(password, target=MAIL_CRED_TARGET, comment="SMTP mailbox password for solar invoices")
                self.respond(200, {"saved": True})
                return
            if parsed.path == "/api/send-invoice":
                relative = str(payload.get("path", "")).replace("\\", "/")
                history_path = OUTPUT / "history.json"
                entries = load_history()
                entry = next((x for x in entries if x.get("path") == relative), None)
                if entry is None:
                    raise RuntimeError("Rechnung nicht im aktiven Register gefunden.")
                pdf_path = send_invoice_email(entry, str(payload.get("recipient", "")))
                entry["email_sent_to"] = email.utils.parseaddr(str(payload.get("recipient", "")))[1]
                entry["email_sent_at"] = datetime.now(ZoneInfo("Europe/Zurich")).isoformat(timespec="minutes")
                s.history_index(entries)
                self.respond(200, {"sent": True, "pdf": pdf_path.relative_to(OUTPUT).as_posix()})
                return
            if parsed.path == "/api/settings":
                self.respond(200, save_settings(payload))
                return
            if parsed.path == "/api/delete":
                relative = str(payload.get("path", "")).replace("\\", "/")
                target = (OUTPUT / relative).resolve()
                if not target.is_relative_to(OUTPUT.resolve()) or target.suffix.lower() != ".html" or not target.is_file():
                    raise RuntimeError("Rechnungsdatei nicht gefunden oder ungültiger Pfad.")
                history_path = OUTPUT / "history.json"
                entries = load_history()
                entry = next((x for x in entries if x.get("path") == relative), None)
                if entry is None:
                    raise RuntimeError("Diese Rechnung ist nicht mehr im Register.")
                trash = OUTPUT / "Papierkorb"
                trash.mkdir(parents=True, exist_ok=True)
                destination = trash / target.name
                if destination.exists():
                    destination = trash / (target.stem + "-" + str(int(time.time())) + target.suffix)
                shutil.move(str(target), str(destination))
                entries = [x for x in entries if x.get("path") != relative]
                s.history_index(entries)
                self.respond(200, {"deleted": True, "recoverable_path": destination.relative_to(OUTPUT).as_posix()})
                return
            if parsed.path != "/api/generate":
                self.respond(404, {"error": "Nicht gefunden"})
                return
            result = generate(payload.get("files", []), (payload.get("apiKey") or "").strip(), payload.get("settings", {}))
            self.respond(200, result)
        except Exception as exc:
            self.respond(400, {"error": str(exc)})


def html_to_pdf(invoice_path):
    candidates = [
        Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
        Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
        Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
        Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
    ]
    browser = next((p for p in candidates if p.is_file()), None)
    if browser is None:
        raise RuntimeError("Für den PDF-Anhang wurde Edge oder Chrome nicht gefunden.")
    pdf_path = invoice_path.with_suffix(".pdf")
    profile = Path(tempfile.mkdtemp(prefix="solar-pdf-profile-"))
    try:
        cmd = [
            str(browser), "--headless", "--disable-gpu", "--no-first-run",
            "--no-default-browser-check", f"--user-data-dir={profile}",
            "--no-pdf-header-footer", f"--print-to-pdf={pdf_path}", invoice_path.as_uri(),
        ]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
        if proc.returncode != 0 or not pdf_path.is_file() or pdf_path.stat().st_size < 1000:
            detail = (proc.stderr or proc.stdout or "").strip()[-500:]
            raise RuntimeError("PDF konnte nicht erstellt werden. " + detail)
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    return pdf_path


def send_invoice_email(entry, recipient):
    parsed_name, parsed_email = email.utils.parseaddr(recipient.strip())
    if not parsed_email or "@" not in parsed_email or "." not in parsed_email.rsplit("@", 1)[-1]:
        raise RuntimeError("Bitte eine gültige Empfänger-E-Mail-Adresse eingeben.")
    if not str(s.CFG.get("email", "")).strip() or not str(s.CFG.get("smtp_host", "")).strip():
        raise RuntimeError("Bitte Absender-E-Mail und SMTP-Server in abrechnungskonfiguration.json eintragen.")
    password = s.cred_read(MAIL_CRED_TARGET)
    if not password:
        raise RuntimeError("Bitte zuerst das SMTP-Mailbox-Passwort im Bereich E-Mail-Versand speichern.")
    invoice_path = (OUTPUT / entry["path"]).resolve()
    if not invoice_path.is_relative_to(OUTPUT.resolve()) or not invoice_path.is_file():
        raise RuntimeError("Die gespeicherte Rechnung wurde nicht gefunden.")
    pdf_path = html_to_pdf(invoice_path)
    msg = EmailMessage()
    msg["From"] = s.CFG["email"]
    msg["To"] = parsed_email
    msg["Subject"] = f'Stromabrechnung {entry["period"]} – Rechnung {entry["invoice"]}'
    msg.set_content(
        f'Guten Tag {entry["customer"]},\n\n'
        f'anbei erhalten Sie die Stromabrechnung {entry["period"]} '
        f'als PDF. Rechnungsbetrag: CHF {float(entry["amount"]):.2f}.\n\n'
        "Freundliche Grüsse\n" + str(s.CFG["sender"])
    )
    msg.add_attachment(pdf_path.read_bytes(), maintype="application", subtype="pdf", filename=pdf_path.name)
    try:
        smtp_host = str(s.CFG.get("smtp_host", "")).strip()
        smtp_port = int(s.CFG.get("smtp_port", 465))
        if not smtp_host:
            raise RuntimeError("Bitte den SMTP-Server in abrechnungskonfiguration.json eintragen.")
        with smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=45, context=ssl.create_default_context()) as smtp:
            smtp.login(s.CFG["email"], password)
            smtp.send_message(msg)
    except smtplib.SMTPAuthenticationError as exc:
        raise RuntimeError("Der SMTP-Server hat die Anmeldung abgelehnt. Prüfe Server, Port, Absenderadresse und Mailbox-Passwort.") from exc
    except Exception as exc:
        raise RuntimeError(f"SMTP-Versand fehlgeschlagen: {exc}") from exc
    return pdf_path

def main():
    try:
        server = ThreadingHTTPServer((HOST, PORT), Handler)
    except OSError:
        return
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()





























