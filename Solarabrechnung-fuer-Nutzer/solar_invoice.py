# -*- coding: utf-8 -*-
import ctypes, ctypes.wintypes, html, json, re, sys, urllib.parse, urllib.request, webbrowser
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from openpyxl import load_workbook
from reportlab.graphics import renderSVG
from reportlab.graphics.barcode.qr import QrCodeWidget
from reportlab.graphics.shapes import Drawing
from reportlab.lib.units import mm

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "Ausgabe"
API = "https://cloud.solar-manager.ch"
CRED_TARGET = "Solarabrechnung.SolarManagerApiKey"
CONFIG_FILE = ROOT / "abrechnungskonfiguration.json"
DEFAULT_CFG = {
    "iban":"", "sender":"", "street":"", "house":"", "postal":"", "city":"", "country":"CH",
    "email":"", "phone":"", "smtp_host":"", "smtp_port":465,
    "grid_rp":0, "solar_rp":0, "meter_rent_chf":0, "due_days":30, "sm_id":"", "customers":{}
}

def load_config():
    if not CONFIG_FILE.exists():
        return DEFAULT_CFG.copy()
    data=json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    if not isinstance(data,dict):
        raise RuntimeError("abrechnungskonfiguration.json muss ein JSON-Objekt enthalten.")
    cfg=DEFAULT_CFG.copy(); cfg.update(data)
    return cfg

CFG = load_config()

def save_profile(data):
    fields = ("iban", "sender", "street", "house", "postal", "city")
    values = {name: str(data.get(name, "")).strip() for name in fields}
    missing = [name for name, value in values.items() if not value]
    if missing:
        raise RuntimeError("Bitte alle Absender- und Zahlungsfelder ausfüllen: " + ", ".join(missing))
    updated = CFG.copy()
    updated.update(values)
    CONFIG_FILE.write_text(json.dumps(updated, ensure_ascii=False, indent=2), encoding="utf-8")
    CFG.update(values)
    return {name: CFG[name] for name in fields}

def save_customers(rows):
    if not isinstance(rows, list) or not rows:
        raise RuntimeError("Bitte mindestens einen Bezüger erfassen.")
    customers = {}
    for row in rows:
        meter_id = re.sub(r"\D", "", str(row.get("meterId", "")))
        if len(meter_id) < 5:
            raise RuntimeError("Jede Zählernummer muss mindestens fünf Ziffern enthalten.")
        if meter_id in customers:
            raise RuntimeError(f"Die Zählernummer {meter_id} ist doppelt eingetragen.")
        required = ("name", "street", "house", "postal", "city")
        customer = {key: str(row.get(key, "")).strip() for key in required}
        missing = [key for key, value in customer.items() if not value]
        if missing:
            raise RuntimeError(f"Bei Zähler {meter_id} fehlen Angaben: " + ", ".join(missing))
        customer.update({
            "email": str(row.get("email", "")).strip(),
            "meter": str(row.get("meter", "")).strip() or f"Zähler {meter_id}",
            "country": "CH",
        })
        customers[meter_id] = customer
    updated = CFG.copy()
    updated["customers"] = customers
    CONFIG_FILE.write_text(json.dumps(updated, ensure_ascii=False, indent=2), encoding="utf-8")
    CFG["customers"] = customers
    return [{"meterId": meter_id, **customer} for meter_id, customer in customers.items()]

class CREDENTIALW(ctypes.Structure):
    _fields_=[("Flags",ctypes.wintypes.DWORD),("Type",ctypes.wintypes.DWORD),("TargetName",ctypes.wintypes.LPWSTR),("Comment",ctypes.wintypes.LPWSTR),("LastWritten",ctypes.wintypes.FILETIME),("CredentialBlobSize",ctypes.wintypes.DWORD),("CredentialBlob",ctypes.POINTER(ctypes.c_ubyte)),("Persist",ctypes.wintypes.DWORD),("AttributeCount",ctypes.wintypes.DWORD),("Attributes",ctypes.c_void_p),("TargetAlias",ctypes.wintypes.LPWSTR),("UserName",ctypes.wintypes.LPWSTR)]

def cred_read(target=CRED_TARGET):
    ptr=ctypes.POINTER(CREDENTIALW)()
    if not ctypes.windll.advapi32.CredReadW(target,1,0,ctypes.byref(ptr)): return None
    try: return ctypes.string_at(ptr.contents.CredentialBlob,ptr.contents.CredentialBlobSize).decode("utf-16-le")
    finally: ctypes.windll.advapi32.CredFree(ptr)

def cred_write(secret,target=CRED_TARGET,comment="Solar Manager credentials"):
    raw=secret.encode("utf-16-le"); buf=ctypes.create_string_buffer(raw); c=CREDENTIALW()
    c.Type=1; c.TargetName=target; c.Comment=comment
    c.CredentialBlobSize=len(raw); c.CredentialBlob=ctypes.cast(buf,ctypes.POINTER(ctypes.c_ubyte)); c.Persist=2; c.UserName="Solar Manager API"
    if not ctypes.windll.advapi32.CredWriteW(ctypes.byref(c),0): raise ctypes.WinError()

def get_key():
    key=cred_read()
    if key: return key
    root=tk.Tk(); root.withdraw(); root.attributes("-topmost",True)
    key=simpledialog.askstring("Solar Manager verbinden","API-Schlüssel eingeben. Er wird verschlüsselt lokal gespeichert:",show="*",parent=root)
    root.destroy()
    key=(key or "").strip()
    if not key: raise RuntimeError("Kein API-Schlüssel eingegeben.")
    cred_write(key)
    return key
def call(url, data=None, token=None):
    headers={"Accept":"application/json"}
    if data is not None: headers["Content-Type"]="application/json"
    if token: headers["Authorization"]="Bearer "+token
    body=json.dumps(data).encode() if data is not None else None
    req=urllib.request.Request(url,data=body,headers=headers,method="POST" if body else "GET")
    try:
        with urllib.request.urlopen(req,timeout=45) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail=e.read().decode("utf-8","replace").strip()
        route=urllib.parse.urlparse(url).path
        suffix=f": {detail[:400]}" if detail else ""
        raise RuntimeError(f"Solar-Manager-API {route} verweigert den Zugriff (HTTP {e.code}){suffix}") from e
    except Exception as e:
        route=urllib.parse.urlparse(url).path
        raise RuntimeError(f"Solar-Manager-Verbindung zu {route} fehlgeschlagen: {e}") from e

def api_token():
    key=get_key()
    x=call(API+"/v3/auth/refresh",{"grant_type":"refresh_token","refresh_token":key})
    if not x.get("access_token"): raise RuntimeError("Solar Manager hat kein Zugriffstoken zurückgegeben. Prüfe den neu erzeugten Cloud-API-Schlüssel.")
    return x["access_token"]

def gateway(token):
    if CFG["sm_id"]: return CFG["sm_id"]
    a=call(API+"/v1/users",token=token)
    if isinstance(a,dict): a=a.get("data",a.get("users",[]))
    choices=[u for u in a if isinstance(u,dict) and u.get("sm_id")]
    if len(choices)==1: return choices[0]["sm_id"]
    if not choices: raise RuntimeError("Keine Solar-Manager-ID (sm_id) gefunden.")
    labels=[f'{u.get("plant") or (u.get("first_name","")+" "+u.get("last_name",""))} [{u["sm_id"]}]' for u in choices]
    prompt="Anlagennummer: "+"; ".join(f"{i+1}. {label}" for i,label in enumerate(labels))
    root=tk.Tk(); root.withdraw(); root.attributes("-topmost",True)
    choice=simpledialog.askinteger("Solar Manager Anlage",prompt,minvalue=1,maxvalue=len(choices),parent=root)
    root.destroy()
    if not choice: raise RuntimeError("Keine Solar-Manager-Anlage ausgewählt.")
    return choices[choice-1]["sm_id"]
def num(v):
    if isinstance(v,(int,float)): return float(v)
    s=str(v).strip().replace(" ","").replace("'","")
    if "," in s and "." in s:
        s=s.replace(".","").replace(",",".") if s.rfind(",")>s.rfind(".") else s.replace(",","")
    else: s=s.replace(",",".")
    return float(s)

def read_export(path):
    wb=load_workbook(path,data_only=True,read_only=True)
    try:
        ws=next((s for s in wb.worksheets if re.search(r"\d{5,}",str(s["B2"].value or "")) and s["B3"].value),None)
        if ws is None:
            raise RuntimeError(f"{path.name}: Smart-me-Blatt mit Verbraucher in B2 und Zeitraum in B3 fehlt.")
        meter=re.search(r"\d{5,}",str(ws["B2"].value)).group()
        if meter not in CFG["customers"]:
            raise RuntimeError(f"Zähler {meter} fehlt in der Bezügerliste.")
        dates=re.findall(r"\d{2}\.\d{2}\.\d{4}",str(ws["B3"].value))
        if len(dates)<2:
            raise RuntimeError(f"Zeitraum nicht lesbar: {ws['B3'].value}")
        start=datetime.strptime(dates[0],"%d.%m.%Y").date()
        end=datetime.strptime(dates[1],"%d.%m.%Y").date()
        energy=None
        for row in ws.iter_rows(min_row=1,values_only=True):
            label=str(row[0] or "").strip().casefold() if row else ""
            if label in ("import total","bezug total"):
                try:
                    energy=num(row[4])
                    break
                except (TypeError,ValueError,IndexError):
                    pass
        if energy is None:
            raise RuntimeError(f"Gesamtbezug in Spalte E fehlt: {path.name}")
        return {"meter":meter,"start":start,"end":end,"kwh":energy,"source":path.name}
    finally:
        wb.close()
def quarter(start,end):
    q=(start.month-1)//3+1; first=date(start.year,(q-1)*3+1,1)
    nxt=(q-1)*3+4; nxtdate=date(start.year+(nxt>12),(nxt-1)%12+1,1); last=nxtdate-timedelta(days=1)
    if start!=first or end!=last: raise RuntimeError(f"Exportzeitraum {start:%d.%m.%Y}–{end:%d.%m.%Y} ist kein volles Kalenderquartal.")
    return start.year,q

def utc_iso(d,last=False):
    v=datetime.combine(d,time.max if last else time.min,tzinfo=ZoneInfo("Europe/Zurich")).astimezone(timezone.utc)
    return v.isoformat(timespec="milliseconds").replace("+00:00","Z")

def stats(token,smid,start,end):
    q=urllib.parse.urlencode({"accuracy":"low","from":utc_iso(start),"to":utc_iso(end,True)})
    s=call(f"{API}/v1/statistics/gateways/{urllib.parse.quote(smid,safe='')}?{q}",token=token)
    for k in ("production","selfConsumptionRate","autarchyDegree"):
        if k not in s: raise RuntimeError("Solar-Manager-Antwort unvollständig.")
    return s

def ref(number):
    digits="".join(x for x in number if x.isdigit())[-21:] or "1"
    check=98-(int(digits+"271500")%97)
    return f"RF{check:02d}{digits}"

def qr(invoice):
    c=CFG; d=invoice["customer"]
    lines=["SPC","0200","1",c["iban"],"S",c["sender"],c["street"],c["house"],c["postal"],c["city"],c["country"],
           "","","","","","","",f'{invoice["amount"]:.2f}',"CHF","S",d["name"],d["street"],d["house"],d["postal"],d["city"],d["country"],
           "SCOR",invoice["ref"],"Stromabrechnung "+invoice["period"],"EPD","",""]
    payload="\n".join(lines)
    widget=QrCodeWidget(payload,barLevel="M",barWidth=46*mm,barHeight=46*mm,barBorder=0)
    drawing=Drawing(46*mm,46*mm); drawing.add(widget)
    svg=renderSVG.drawToString(drawing)
    logo='<svg x="42.3913%" y="42.3913%" width="15.2174%" height="15.2174%" viewBox="0 0 7 7"><rect width="7" height="7" fill="white" stroke="black" stroke-width=".12"/><rect x=".5" y=".5" width="6" height="6" fill="black"/><path d="M3 1.4h1v2.1h2.1v1H4v2.1H3V4.5H.9v-1H3z" fill="white"/></svg>'
    return svg.replace("</svg>",logo+"</svg>")

def chf(v): return f"{v:,.2f}".replace(",","X").replace(".",",").replace("X","’")
def kwh(v): return f"{v:,.2f}".replace(",","X").replace(".",",").replace("X","’")
def mwh(v): return f"{v/1000:,.2f}".replace(",","X").replace(".",",").replace("X","’")

def render(inv,solar):
    c=CFG; d=inv["customer"]; qrcode=qr(inv); prod=float(solar["production"])/1000; total_consumption_kwh=float(solar["consumption"])/1000; aut_kwh=float(solar["selfConsumption"])/1000; self_kwh=float(solar["selfConsumption"])/1000
    monthly = solar.get("monthlyProduction", [])
    monthly_rows = "".join(f'<div><span>{html.escape(str(entry["month"]))}</span><b>{mwh(float(entry["production"]))} MWh</b></div>' for entry in monthly)
    monthly_production = f'<section class="monthly"><h2>Monatliche Produktion der Gesamtanlage</h2>{monthly_rows}</section>' if monthly else ""
    exported_kwh=max(0.0,prod-self_kwh)
    self_share_pct=min(100.0,max(0.0,self_kwh/prod*100)) if prod>0 else 0.0
    exported_share_pct=max(0.0,100.0-self_share_pct)
    self_share_display=f"{self_share_pct:.1f}".replace(".",",")
    exported_share_display=f"{exported_share_pct:.1f}".replace(".",",")
    production_chart=f'<section class="usage-chart"><h3>Verwendung der Solarproduktion der Gesamtanlage</h3><div class="usage-bar"><span style="width:{self_share_pct:.1f}%"></span></div><div class="usage-labels"><span>Selbst genutzt: {kwh(self_kwh)} kWh ({self_share_display} %)</span><span>Nicht selbst genutzt: {kwh(exported_kwh)} kWh ({exported_share_display} %)</span></div></section>'
    grid_rp_value=float(inv.get("grid_rp",c["grid_rp"])); solar_rp_value=float(inv.get("solar_rp",c["solar_rp"])); gp=grid_rp_value/100; sp=solar_rp_value/100; esc=html.escape
    saving_pct=max(0,(grid_rp_value-solar_rp_value)/grid_rp_value*100) if grid_rp_value>0 else 0
    autarchy_pct=float(solar["autarchyDegree"])
    allocation_pct=float(inv.get("solar_allocation_pct",autarchy_pct))
    allocation_display=f"{allocation_pct:.1f}".replace(".",",")
    grid_display=f"{100-allocation_pct:.1f}".replace(".",",")
    saving_display=f"{saving_pct:.1f}".replace(".",",")
    solar_desc=f'<div class="share-values"><span>Anteil am Verbrauch</span><span>Solar <b>{allocation_display} %</b> · Netz <b>{grid_display} %</b></span></div><div class="share-bar"><i style="width:{allocation_pct:.1f}%"></i></div><div class="share-saving">Solarstrom {saving_display} % günstiger als Netzstrom</div>'
    reduction=float(inv.get("reduction_amount",0)); subtotal=float(inv.get("subtotal",inv["grid_amt"]+inv["solar_amt"])); meter_rent=float(inv.get("meter_rent",c.get("meter_rent_chf",8.00)))
    subtotal_row=f'<tr class="subtotal"><td colspan="3">Zwischentotal</td><td class="right">CHF {chf(subtotal)}</td></tr>' if reduction>0 else ""
    discount_row=f'<tr class="discount"><td colspan="3">{esc(inv.get("reduction_description","Reduktion"))}</td><td class="right">− CHF {chf(reduction)}</td></tr>' if reduction>0 else ""
    rent_row=f'<tr class="fee"><td colspan="3">Zählermiete</td><td class="right">CHF {chf(meter_rent)}</td></tr>' if meter_rent>0 else ""
    return f"""<!doctype html><html lang="de-CH"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Stromrechnung {esc(inv["no"])}</title><style>
:root{{--ink:#183137;--muted:#63777a;--line:#dce6e3;--green:#176b5b;--mint:#edf6f2;--gold:#c39b55}}*{{box-sizing:border-box}}body{{margin:0;background:#eef2f0;color:var(--ink);font:13px/1.42 Arial,sans-serif}}.page{{width:210mm;height:297mm;margin:16px auto;background:#fff;box-shadow:0 8px 28px #18313720}}.main{{height:215mm;padding:14mm 16mm 7mm}}header{{display:flex;justify-content:space-between;border-bottom:1px solid var(--line);padding-bottom:5mm}}.brand{{display:flex;gap:11px;align-items:center}}.mark{{width:46px;height:46px;border-radius:14px;background:linear-gradient(145deg,#174d44,#21806a);display:grid;place-items:center;overflow:hidden;box-shadow:0 5px 13px #176b5b30;flex:none}}.mark svg{{width:46px;height:46px;display:block}}.name{{font-size:12px;font-weight:bold;letter-spacing:.09em;text-transform:uppercase}}.sub{{color:var(--muted);font-size:10px;margin-top:3px}}.title{{text-align:right}}.title small,.label{{color:var(--muted);font-size:9px;text-transform:uppercase;letter-spacing:.1em}}h1{{font:normal 30px Georgia,serif;margin:2px 0}}.meta{{display:grid;grid-template-columns:repeat(3,1fr);margin:3mm 0;gap:12px}}.meta div{{border-left:2px solid var(--gold);padding-left:9px}}.meta b{{display:block;font-size:12px;margin-top:1px}}.parties{{display:grid;grid-template-columns:1fr 1fr;gap:16mm;margin-bottom:3mm}}.party b{{display:block;font-size:14px;margin:3px 0 2px}}.party p{{margin:0;color:#435a5e;font-size:11px;line-height:1.55}}.meter{{color:var(--muted)!important;margin-top:5px!important}}.period{{background:var(--mint);border:1px solid #dcece5;border-radius:9px;padding:7px 10px;display:flex;justify-content:space-between;margin-bottom:3mm}}.period b{{font-size:12px}}h2{{font-size:11px;letter-spacing:.04em;margin:0 0 4px}}table{{width:100%;border-collapse:collapse}}th{{text-align:left;font-size:8px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;border-bottom:1px solid var(--line);padding:8px 5px}}td{{font-size:10px;border-bottom:1px solid var(--line);padding:6px 5px}}td:last-child,th:last-child{{text-align:right}}.right{{text-align:right;white-space:nowrap}}.desc b{{display:block}}.desc small{{color:var(--muted)}}.share-bar{{height:3px;max-width:65mm;background:#dce6e3;border-radius:4px;overflow:hidden;margin:2px 0}}.share-bar i{{display:block;height:100%;background:#176b5b}}.share-values{{display:flex;justify-content:space-between;max-width:92mm;gap:4px;font-size:8px;color:#52676b}}.share-values b{{display:inline!important}}.share-saving{{font-size:8px;color:#176b5b;font-weight:bold;margin-top:2px}}.subtotal td{{border-top:1px solid var(--line);padding:8px 5px 3px;text-align:right;font-size:10px;color:var(--muted)}}.discount td{{padding:4px 5px;color:#a34d30;font-size:10px;text-align:right}}.fee td{{padding:6px 5px;font-size:10px;border-top:1px solid var(--line)}}.total td{{border:0;padding-top:12px;font-size:13px;font-weight:bold}}.total td:last-child{{color:var(--green);font-size:17px}}.saving{{text-align:right;color:var(--green);font-size:9px;margin:0 4px 1mm}}.stats{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-top:2mm}}.usage-chart{{margin-top:1.5mm;padding:3px 7px;border:1px solid #dce6e3;border-radius:7px;background:#fbfdfb}}.usage-chart h3{{font-size:8px;margin:0 0 2px;color:#52676b}}.usage-bar{{height:4px;border-radius:5px;background:#c39b55;overflow:hidden}}.usage-bar span{{display:block;height:100%;background:#176b5b}}.usage-labels{{display:flex;justify-content:space-between;gap:8px;font-size:7px;color:#52676b;margin-top:2px}}.stat{{border:1px solid var(--line);border-radius:8px;padding:5px 8px;min-height:14mm}}.stat:nth-child(1){{background:#e8f3ef;border-color:#b9d9cc}}.stat:nth-child(2){{background:#eaf0fb;border-color:#c3d4f2}}.stat:nth-child(3){{background:#fff4e5;border-color:#f0d7a5}}.stat span{{display:block;color:#52676b;font-size:8px;text-transform:uppercase;letter-spacing:.07em}}.stat:nth-child(1) b{{color:#176b5b}}.stat:nth-child(2) b{{color:#335d9a}}.stat:nth-child(3) b{{color:#93621d}}.stat b{{display:block;font-size:15px;margin-top:4px}}.stat small{{display:block;font-size:9px;color:#52676b;margin-top:2px}}.monthly{{margin-top:1mm;border:1px solid #b9d9cc;background:#f3f8f5;border-radius:8px;padding:3px 8px;display:grid;grid-template-columns:repeat(3,1fr);gap:1px 10px}}.monthly h2{{grid-column:1/-1;margin:0;font-size:8px;letter-spacing:0}}.monthly div{{display:flex;justify-content:space-between;font-size:8px}}.foot{{font-size:8px;line-height:1.35;color:var(--muted);margin-top:2mm;margin-bottom:1mm;max-width:165mm}}.contact{{border-top:1px solid var(--line);padding-top:2mm;margin-top:4mm;display:flex;justify-content:space-between;font-size:8px;color:var(--muted)}}.slip{{height:82mm;border-top:1px dashed #8c9996;display:grid;grid-template-columns:62mm 148mm;position:relative;font-size:8px;line-height:1.3}}.receipt{{padding:4mm;border-right:1px dashed #8c9996}}.payment{{padding:4mm 7mm;position:relative}}.slip h3{{font-size:10px;margin:0 0 2mm}}.slab{{font-size:7px;color:var(--muted);margin:0 0 1mm}}.sv{{font-size:8px;font-weight:600;margin:0 0 2mm;line-height:1.3}}.ref{{font-family:monospace;overflow-wrap:anywhere}}.amount{{display:flex;gap:10mm;margin:3mm 0}}.paygrid{{display:grid;grid-template-columns:56mm 1fr;gap:5mm}}.qr{{width:56mm;height:56mm;display:flex;align-items:center;justify-content:center}}.qr svg{{width:46mm;height:46mm}}.cut{{position:absolute;top:-4mm;left:50%;transform:translateX(-50%);background:#fff;color:var(--muted);font-size:8px;padding:0 4px}}@media print{{@page{{size:A4;margin:0}}html,body{{width:210mm;height:297mm;margin:0;background:white}}.print-btn{{display:none}}.page{{position:relative;width:210mm;height:297mm;margin:0;box-shadow:none;overflow:hidden;page-break-after:always;break-after:page}}.main{{position:absolute;left:0;top:0;width:210mm;height:215mm;padding:7mm 16mm 3mm;overflow:hidden}}.slip{{position:absolute;left:0;bottom:0;width:210mm;height:82mm;display:grid;grid-template-columns:62mm 148mm;overflow:hidden}}.payment .qr{{position:absolute;left:7mm;bottom:-4mm;top:auto}}.payment .paygrid>div:first-child .slab{{position:absolute;left:7mm;bottom:50mm}}}}@media screen and (max-width:760px){{body{{background:white}}.page{{width:100%;height:auto;margin:0;box-shadow:none}}.main{{height:auto;padding:22px}}.slip{{height:auto;min-height:350px;grid-template-columns:1fr 2fr}}.paygrid{{grid-template-columns:1fr}}.qr{{justify-content:flex-start}}}}</style></head><body><button class="print-btn" onclick="window.print()" style="position:fixed;z-index:10;top:12px;right:14px;background:#176b5b;color:white;border:0;border-radius:10px;padding:10px 14px;font:700 12px Arial;cursor:pointer">Als PDF speichern</button><section class="page"><div class="main">
<header><div class="brand"><div class="mark"><svg viewBox="0 0 64 64" role="img" aria-label="Solarenergie"><circle cx="47" cy="15" r="7" fill="#e6bf70"/><path d="M10 34 31.5 18 54 34" fill="none" stroke="#f8f5e9" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/><path d="m17 34 28 0 4 16H13z" fill="#d8eee4"/><path d="M27 35 24 49M36 35 37 49M15 42h32" stroke="#31806d" stroke-width="1.5"/><path d="M22 55h20" stroke="#e6bf70" stroke-width="2" stroke-linecap="round"/><text x="32" y="62" fill="#f8f5e9" font-family="Arial,sans-serif" font-size="7" font-weight="700" text-anchor="middle" letter-spacing="1">SA</text></svg></div><div><div class="name">{esc(c["sender"])}</div><div class="sub">{esc(c["street"])} {c["house"]} · {c["postal"]} {esc(c["city"])}</div></div></div><div class="title"><small>Stromabrechnung</small><h1>Rechnung</h1></div></header>
<div class="meta"><div><span class="label">Rechnungsnummer</span><b>{esc(inv["no"])}</b></div><div><span class="label">Rechnungsdatum</span><b>{inv["date"]:%d.%m.%Y}</b></div><div><span class="label">Zahlbar bis</span><b>{inv["due"]:%d.%m.%Y}</b></div></div>
<div class="parties"><div class="party"><span class="label">Rechnung an</span><b>{esc(d["name"])}</b><p>{esc(d["street"])} {d["house"]}<br>{d["postal"]} {esc(d["city"])}</p><p class="meter">{esc(d["meter"])}</p></div><div class="party"><span class="label">Rechnung von</span><b>{esc(c["sender"])}</b><p>{esc(c["street"])} {c["house"]}<br>{c["postal"]} {esc(c["city"])}</p><p class="meter">{esc(c["email"])} · {esc(c["phone"])}</p></div></div>
<div class="period"><div><span class="label">Abrechnungszeitraum</span><br><b>{inv["start"]:%d.%m.%Y} – {inv["end"]:%d.%m.%Y}</b></div><div style="text-align:right"><span class="label">Strombezug</span><br><b>{kwh(inv["cons"])} kWh</b></div></div>
<h2>Abgerechnete Energie</h2><table><thead><tr><th>Leistung</th><th class="right">Menge</th><th class="right">Preis je kWh</th><th class="right">Betrag</th></tr></thead><tbody><tr><td class="desc"><b>Netzstrom</b><small>Gesamtbezug abzüglich Solarstromanteil</small></td><td class="right">{kwh(inv["grid"])} kWh</td><td class="right">CHF {gp:.4f}</td><td class="right">CHF {chf(inv["grid_amt"])}</td></tr><tr><td class="desc"><b>Solarstrom</b><div class="solar-share">{solar_desc}</div></td><td class="right">{kwh(inv["solar"])} kWh</td><td class="right">CHF {sp:.4f}</td><td class="right">CHF {chf(inv["solar_amt"])}</td></tr>{subtotal_row}{discount_row}{rent_row}<tr class="total"><td colspan="3">Zu bezahlen</td><td>CHF {chf(inv["amount"])}</td></tr></tbody></table>
<p class="saving">Vorteil durch Solarstrom: CHF {chf(inv["savings"])}</p><p class="foot">Anlagenwerte für den Abrechnungszeitraum; oben wird der Anteil dieses Bezügers abgerechnet.</p><h2 style="margin-top:1mm">Gesamte Solaranlage · {esc(inv["period"])}</h2><div class="stats"><div class="stat"><span>Autarkiegrad der Anlage</span><b>{solar["autarchyDegree"]:.2f} %</b><small>{kwh(aut_kwh)} von {kwh(total_consumption_kwh)} kWh Gesamtverbrauch</small></div><div class="stat"><span>Eigenverbrauchsanteil der Produktion</span><b>{solar["selfConsumptionRate"]:.2f} %</b><small>{kwh(self_kwh)} von {kwh(prod)} kWh Solarproduktion</small></div><div class="stat"><span>Produktion der Gesamtanlage</span><b>{kwh(prod)} kWh</b></div></div>
{production_chart}{monthly_production}<div class="contact"><span>Vielen Dank.</span><span>{esc(c["sender"])} · IBAN {esc(c["iban"])}</span></div>
</div><div class="slip"><span class="cut">✂</span><div class="receipt"><h3>Empfangsschein</h3><p class="slab">Konto / Zahlbar an</p><p class="sv">{esc(c["iban"])}<br>{esc(c["sender"])}<br>{esc(c["street"])} {c["house"]}<br>{c["postal"]} {esc(c["city"])}</p><p class="slab">Referenz</p><p class="sv ref">{esc(inv["ref"])}</p><div class="amount"><div><p class="slab">Währung</p><b>CHF</b></div><div><p class="slab">Betrag</p><b>{chf(inv["amount"])}</b></div></div><p class="slab">Zahlbar durch</p><p class="sv">{esc(d["name"])}<br>{esc(d["street"])} {d["house"]}<br>{d["postal"]} {esc(d["city"])}</p></div>
<div class="payment"><h3>Zahlteil</h3><div class="paygrid"><div><div class="qr">{qrcode}</div><p class="slab">Mit der Banking-App scannen</p></div><div><p class="slab">Konto / Zahlbar an</p><p class="sv">{esc(c["iban"])}<br>{esc(c["sender"])}<br>{esc(c["street"])} {c["house"]}<br>{c["postal"]} {esc(c["city"])}</p><p class="slab">Referenz</p><p class="sv ref">{esc(inv["ref"])}</p><p class="slab">Zusätzliche Informationen</p><p class="sv">Stromabrechnung {esc(inv["period"])}</p><div class="amount"><div><p class="slab">Währung</p><b>CHF</b></div><div><p class="slab">Betrag</p><b>{chf(inv["amount"])}</b></div></div><p class="slab">Zahlbar durch</p><p class="sv">{esc(d["name"])} · {esc(d["street"])} {d["house"]} · {d["postal"]} {esc(d["city"])}</p></div></div></div></div></section></body></html>"""

def choose_files():
    root=tk.Tk(); root.withdraw(); root.attributes("-topmost",True)
    files=filedialog.askopenfilenames(title="Smart-me-Quartalsexporte auswählen",filetypes=[("Excel-Dateien","*.xlsx")])
    root.destroy(); return [Path(x) for x in files]

def history_index(entries):
    OUTPUT.mkdir(parents=True,exist_ok=True)
    (OUTPUT/"history.json").write_text(json.dumps(entries,ensure_ascii=False,indent=2),encoding="utf-8")
    groups={}
    for entry in entries:
        groups.setdefault(entry["period"],[]).append(entry)
    def period_key(label):
        match=re.search(r"(\d{4})\s+Q([1-4])",label)
        return (int(match.group(1)),int(match.group(2))) if match else (0,0)
    sections=[]
    for period in sorted(groups,key=period_key,reverse=True):
        rows=[]
        for e in sorted(groups[period],key=lambda x:(x["issued"],x["invoice"]),reverse=True):
            href=html.escape(e["path"].replace(chr(92),"/"),quote=True)
            rows.append(f'<tr><td>{html.escape(e["customer"])}</td><td>{html.escape(e["invoice"])}</td><td>{html.escape(e["issued"])}</td><td>CHF {e["amount"]:.2f}</td><td><a href="{href}">Rechnung öffnen</a></td></tr>')
        sections.append(f'<section><h2>{html.escape(period)}</h2><table><thead><tr><th>Bezüger</th><th>Rechnungsnummer</th><th>Ausgestellt</th><th>Betrag</th><th></th></tr></thead><tbody>{"".join(rows)}</tbody></table></section>')
    page="""<!doctype html><html lang="de"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Rechnungshistory</title><style>body{font:15px/1.5 Arial,sans-serif;color:#183137;background:#eef2f0;margin:0;padding:32px}main{max-width:1050px;margin:auto;background:white;padding:28px;border-radius:14px}h1{font:32px Georgia,serif;margin:0 0 6px}h2{font:22px Georgia,serif;color:#176b5b;margin:28px 0 8px}.muted{color:#63777a}table{width:100%;border-collapse:collapse}th,td{text-align:left;padding:12px 10px;border-bottom:1px solid #dce6e3}th{font-size:11px;color:#63777a;text-transform:uppercase;letter-spacing:.06em}a{color:#176b5b;font-weight:bold}.empty{text-align:center;padding:40px;color:#63777a}@media(max-width:700px){body{padding:10px}main{padding:16px;overflow:auto}table{min-width:700px}}</style></head><body><main><h1>Solarstrom-Rechnungen</h1><div class="muted">Archiv · gespeicherte Rechnungen nach Quartal und Bezüger</div>"""
    page+="".join(sections) if sections else '<div class="empty">Noch keine Rechnungen erstellt.</div>'
    page+="</main></body></html>"
    (OUTPUT/"Index.html").write_text(page,encoding="utf-8")
def main():
    files=choose_files()
    if not files:
        messagebox.showinfo("Solarabrechnung","Kein Smart-me-Export ausgewählt.")
        return
    records=[read_export(p) for p in files]; first=records[0]
    y,q=quarter(first["start"],first["end"])
    if any((x["start"],x["end"])!=(first["start"],first["end"]) for x in records):
        raise RuntimeError("Alle Dateien müssen denselben Quartalszeitraum enthalten.")
    if len({x["meter"] for x in records})!=len(records):
        raise RuntimeError("Ein Smart-me-Zähler wurde mehrfach ausgewählt.")
    token=api_token(); smid=gateway(token)
    query=urllib.parse.urlencode({"accuracy":"low","from":utc_iso(first["start"]),"to":utc_iso(first["end"],True)})
    solar=call(f'{API}/v1/statistics/gateways/{urllib.parse.quote(smid,safe="")}?{query}',token=token)
    if not all(k in solar for k in ("production","selfConsumptionRate","autarchyDegree")):
        raise RuntimeError("Solar-Manager-Kennzahlen fehlen.")
    today=datetime.now(ZoneInfo("Europe/Zurich")).date(); period=f"{y} Q{q}"
    folder=OUTPUT/f"{y}-Q{q}"; folder.mkdir(parents=True,exist_ok=True)
    history_path=OUTPUT/"history.json"
    entries=json.loads(history_path.read_text(encoding="utf-8")) if history_path.exists() else []
    used=len(entries)
    for r in records:
        d=CFG["customers"][r["meter"]]
        solar_kwh=min(r["kwh"]*float(solar["autarchyDegree"])/100,r["kwh"])
        grid_kwh=r["kwh"]-solar_kwh
        grid_amt=round(grid_kwh*CFG["grid_rp"]/100,2); solar_amt=round(solar_kwh*CFG["solar_rp"]/100,2)
        amount=round(grid_amt+solar_amt,2); used+=1
        no=f"{y}Q{q}-{used:04d}"
        digits="".join(ch for ch in no if ch.isdigit())[-21:] or "1"
        check=98-(int(digits+"271500")%97); reference=f"RF{check:02d}{digits}"
        inv={"no":no,"ref":reference,"date":today,"due":today+timedelta(days=CFG["due_days"]),"start":r["start"],"end":r["end"],"period":period,"customer":d,"cons":r["kwh"],"solar":solar_kwh,"grid":grid_kwh,"grid_amt":grid_amt,"solar_amt":solar_amt,"amount":amount,"savings":max(0,r["kwh"]*CFG["grid_rp"]/100-amount)}
        path=folder/f"Rechnung-{r['meter']}-{y}-Q{q}-{used:04d}.html"
        path.write_text(render(inv,solar),encoding="utf-8")
        entries.append({"invoice":no,"customer":d["name"],"meter":r["meter"],"period":period,"issued":today.isoformat(),"amount":amount,"path":path.relative_to(OUTPUT).as_posix()})
    history_index(entries)
    webbrowser.open((OUTPUT/"Index.html").as_uri())
    root=tk.Tk(); root.withdraw(); root.attributes("-topmost",True)
    messagebox.showinfo("Solarabrechnung fertig",str(len(records))+" HTML-Rechnung(en) erstellt. Das Rechnungsarchiv wurde im Browser geöffnet. Ausgabeordner: "+str(OUTPUT),parent=root)
    root.destroy()
if __name__=="__main__":
    try:
        main()
    except Exception as e:
        root=tk.Tk(); root.withdraw(); root.attributes("-topmost",True)
        messagebox.showerror("Solarabrechnung – Fehler",str(e),parent=root)
        root.destroy()


























