#!/usr/bin/env python3
"""
Trayectoria EST - CRM Facturación Builder
Lee los Excel de facturación desde SharePoint y genera el HTML del CRM.

Columnas Excel:
  A: Razón Social
  B: RUT
  C: Fecha Factura
  D: Valor Neto
  E: Valor Neto + IVA
  F: N° Factura
  G: Estado
  H: Plazo de Pago
  I: Fecha de Pago
  J: OC
  K: SII
"""

import os
import io
import json
import base64
import urllib.request
import urllib.parse
from datetime import datetime, date
from openpyxl import load_workbook

# ── CONFIGURACIÓN ────────────────────────────────────────────────────────────
SITE        = "https://ssisachile.sharepoint.com/sites/TrayectoriaEST-RemuneracionesyContratos"
DRIVE_ID    = "b!94YSNWupIUmh41_AtdOVSPieLm_WNpBEh9tqCJhq7-HE4RJxxbTATpTpoCXdSMrL"
BASE_PATH   = "Administración y Finanzas/{year}/Contabilidad/Trayectoria EST/Facturación"
FILES       = {
    2026: "Facturación EST 2026.xlsx",
}

# Índices de columna (0-based)
COL_RAZON   = 0   # A - Razón Social
COL_RUT     = 1   # B - RUT
COL_FECHA   = 2   # C - Fecha Factura
COL_NETO    = 3   # D - Valor Neto
COL_TOTAL   = 4   # E - Valor Neto + IVA
COL_NFACT   = 5   # F - N° Factura
COL_ESTADO  = 6   # G - Estado
COL_PLAZO   = 7   # H - Plazo de Pago
COL_FPAGO   = 8   # I - Fecha de Pago
COL_OC      = 9   # J - OC
COL_SII     = 10  # K - SII

# ── AUTH ─────────────────────────────────────────────────────────────────────
def get_token():
    url  = f"https://login.microsoftonline.com/{os.environ['TENANT_ID']}/oauth2/v2.0/token"
    data = urllib.parse.urlencode({
        "grant_type":    "client_credentials",
        "client_id":     os.environ["CLIENT_ID"],
        "client_secret": os.environ["CLIENT_SECRET"],
        "scope":         "https://graph.microsoft.com/.default"
    }).encode()
    req = urllib.request.Request(url, data=data)
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())["access_token"]

# ── SHAREPOINT ────────────────────────────────────────────────────────────────
def get_file_bytes(token, year, filename):
    folder = BASE_PATH.format(year=year)
    path   = f"{folder}/{filename}"
    url    = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/{urllib.parse.quote(path)}:/content"
    req    = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    try:
        with urllib.request.urlopen(req) as r:
            return r.read()
    except Exception as e:
        print(f"⚠️  No se pudo leer {filename}: {e}")
        return None

def patch_cell(token, year, filename, row_number, col_letter, value):
    """Actualiza una celda del Excel en SharePoint via Graph Excel API."""
    folder   = BASE_PATH.format(year=year)
    path     = f"{folder}/{filename}"
    enc_path = urllib.parse.quote(path)
    url      = (f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}"
                f"/root:/{enc_path}:/workbook/worksheets/Sheet1"
                f"/range(address='{col_letter}{row_number}')")
    payload  = json.dumps({"values": [[value]]}).encode()
    req      = urllib.request.Request(
        url, data=payload, method="PATCH",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json"
        }
    )
    with urllib.request.urlopen(req) as r:
        return r.status

# ── PARSEO EXCEL ──────────────────────────────────────────────────────────────
def parse_excel(file_bytes, year):
    wb   = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws   = wb.active
    rows = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(row):
            continue
        def val(idx):
            v = row[idx] if idx < len(row) else None
            return v

        def fmt_date(v):
            if v is None:
                return ""
            if isinstance(v, (datetime, date)):
                return v.strftime("%Y-%m-%d")
            return str(v)

        def fmt_num(v):
            try:
                return float(v) if v not in (None, "") else 0
            except:
                return 0

        rows.append({
            "row":       i,
            "year":      year,
            "razon":     str(val(COL_RAZON) or "").strip(),
            "rut":       str(val(COL_RUT)   or "").strip(),
            "fecha":     fmt_date(val(COL_FECHA)),
            "neto":      fmt_num(val(COL_NETO)),
            "total":     fmt_num(val(COL_TOTAL)),
            "nFactura":  str(val(COL_NFACT) or "").strip(),
            "estado":    str(val(COL_ESTADO) or "").strip(),
            "plazo":     str(val(COL_PLAZO)  or "").strip(),
            "fechaPago": fmt_date(val(COL_FPAGO)),
            "oc":        str(val(COL_OC)     or "").strip(),
            "sii":       str(val(COL_SII)    or "").strip(),
        })
    return rows

# ── BUILD HTML ────────────────────────────────────────────────────────────────
def build_html(facturas):
    today_str = date.today().strftime("%d/%m/%Y")
    fj = json.dumps(facturas, ensure_ascii=False, separators=(',', ':'))

    html = """<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>CRM Facturación — Trayectoria EST</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/react/18.2.0/umd/react.production.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/react-dom/18.2.0/umd/react-dom.production.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/babel-standalone/7.23.2/babel.min.js"></script>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet"/>
  <style>
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}}
    :root{
      --bg:#f0f4fa; --surface:#fff; --card:#fff; --border:#dde4f0; --border-lt:#eef2f9;
      --navy:#1e2d6e; --blue:#3a6fd8; --blue-lt:#5fb3f0; --accent:#3a6fd8; --accent-hv:#2d5bbf; --accent-bg:#eef4fd;
      --danger:#d93b3b; --danger-bg:#fff0f0; --success:#1a9e6a; --success-bg:#edfaf4;
      --warn:#d97b10; --warn-bg:#fff7ed; --text:#1a2340; --text-2:#3d4f72; --muted:#7d90b5;
      --radius:10px; --shadow:0 1px 4px rgba(30,45,110,.08); --shadow-md:0 4px 16px rgba(30,45,110,.10);
    }}
    body{background:var(--bg);color:var(--text);font-family:'Inter',sans-serif;font-size:14px}}
    .crm-root{display:grid;grid-template-columns:240px 1fr;min-height:100vh}}
    .sidebar{background:var(--navy);display:flex;flex-direction:column;position:sticky;top:0;height:100vh}}
    .sidebar-logo{padding:24px 20px 20px;border-bottom:1px solid rgba(255,255,255,.1);margin-bottom:8px}}
    .logo-mark{display:flex;align-items:center;gap:10px}}
    .logo-icon{width:34px;height:34px;flex-shrink:0}}
    .logo-icon svg{width:100%;height:100%}}
    .logo-text{color:#fff;font-size:1.05rem;font-weight:700;line-height:1.2}}
    .logo-text span{display:block;font-size:.7rem;font-weight:400;color:rgba(255,255,255,.5);margin-top:1px}}
    .nav-section{padding:8px 12px;font-size:.65rem;color:rgba(255,255,255,.35);text-transform:uppercase;letter-spacing:1.2px;margin-top:4px}}
    .nav-btn{display:flex;align-items:center;gap:10px;padding:10px 14px;margin:2px 8px;border-radius:8px;border:none;background:transparent;color:rgba(255,255,255,.65);font-family:'Inter',sans-serif;font-size:.84rem;font-weight:500;cursor:pointer;transition:all .15s;text-align:left;width:calc(100% - 16px)}}
    .nav-btn:hover{background:rgba(255,255,255,.08);color:#fff}}
    .nav-btn.active{background:var(--blue);color:#fff;box-shadow:0 2px 8px rgba(58,111,216,.4)}}
    .nav-icon{font-size:.95rem;width:20px;text-align:center;flex-shrink:0}}
    .nav-badge{margin-left:auto;background:var(--danger);color:#fff;font-size:.65rem;font-weight:700;padding:2px 7px;border-radius:20px}}
    .sync-info{margin-top:auto;padding:16px;border-top:1px solid rgba(255,255,255,.1);font-size:.7rem;color:rgba(255,255,255,.35);text-align:center}}
    .main{padding:28px 32px;overflow-y:auto;background:var(--bg)}}
    .page-header{margin-bottom:24px}}
    .page-title{font-size:1.5rem;font-weight:700;color:var(--navy);margin-bottom:4px}}
    .page-sub{color:var(--muted);font-size:.82rem}}
    .stats-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:24px}}
    .stat-card{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);padding:18px 20px;box-shadow:var(--shadow)}}
    .stat-card-top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px}}
    .stat-label{font-size:.72rem;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.7px}}
    .stat-icon{width:34px;height:34px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:1rem}}
    .stat-val{font-size:1.45rem;font-weight:700;color:var(--navy)}}
    .stat-sub{font-size:.72rem;color:var(--muted);margin-top:3px}}
    .table-wrap{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden;margin-bottom:20px;box-shadow:var(--shadow)}}
    .table-header{display:flex;align-items:center;justify-content:space-between;padding:14px 18px;border-bottom:1px solid var(--border-lt);flex-wrap:wrap;gap:10px}}
    .table-title{font-weight:600;font-size:.88rem;color:var(--navy)}}
    table{width:100%;border-collapse:collapse}}
    thead tr{background:#f7f9fd}}
    th{padding:10px 14px;text-align:left;font-size:.7rem;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.8px;white-space:nowrap;border-bottom:1px solid var(--border)}}
    td{padding:11px 14px;font-size:.83rem;border-top:1px solid var(--border-lt);color:var(--text-2)}}
    tr:hover td{background:#f8faff}}
    .badge{display:inline-flex;align-items:center;padding:3px 10px;border-radius:20px;font-size:.7rem;font-weight:600;white-space:nowrap}}
    .badge-danger{background:var(--danger-bg);color:var(--danger)}}
    .badge-warn{background:var(--warn-bg);color:var(--warn)}}
    .badge-success{background:var(--success-bg);color:var(--success)}}
    .badge-blue{background:var(--accent-bg);color:var(--blue)}}
    .badge-muted{background:#f0f2f7;color:var(--muted)}}
    .btn{padding:8px 16px;border-radius:8px;border:none;font-family:'Inter',sans-serif;font-size:.82rem;font-weight:600;cursor:pointer;transition:all .15s;display:inline-flex;align-items:center;gap:6px}}
    .btn-primary{background:var(--accent);color:#fff;box-shadow:0 2px 6px rgba(58,111,216,.25)}}.btn-primary:hover{background:var(--accent-hv)}}
    .btn-ghost{background:transparent;color:var(--text-2);border:1px solid var(--border)}}.btn-ghost:hover{border-color:var(--blue);color:var(--blue)}}
    .btn-success{background:var(--success-bg);color:var(--success);border:1px solid rgba(26,158,106,.2)}}.btn-success:hover{background:#d5f5ea}}
    .btn-warn{background:var(--warn-bg);color:var(--warn);border:1px solid rgba(217,123,16,.2)}}
    .btn-sm{padding:5px 11px;font-size:.76rem}}
    .btn:disabled{opacity:.45;cursor:not-allowed}}
    .search-bar{background:#f5f7fd;border:1px solid var(--border);border-radius:8px;padding:8px 14px;color:var(--text);font-family:'Inter',sans-serif;font-size:.83rem;outline:none;width:220px;transition:border .15s}}
    .search-bar:focus{border-color:var(--blue);background:#fff;box-shadow:0 0 0 3px rgba(58,111,216,.1)}}
    .select-bar{background:#f5f7fd;border:1px solid var(--border);border-radius:8px;padding:8px 14px;color:var(--text);font-family:'Inter',sans-serif;font-size:.83rem;outline:none}}
    .filter-bar{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap;align-items:center}}
    .tabs{display:flex;gap:4px;margin-bottom:20px;background:#f0f4fa;border:1px solid var(--border);border-radius:10px;padding:4px;width:fit-content}}
    .tab{padding:6px 14px;border-radius:7px;border:none;background:transparent;color:var(--muted);font-family:'Inter',sans-serif;font-size:.8rem;font-weight:500;cursor:pointer;transition:all .15s}}
    .tab.active{background:#fff;color:var(--navy);font-weight:700;box-shadow:0 1px 4px rgba(30,45,110,.1)}}
    .empty-state{text-align:center;padding:44px;color:var(--muted)}}
    .progress-bar{height:6px;background:var(--border);border-radius:4px;overflow:hidden;margin-top:4px}}
    .progress-fill{height:100%;border-radius:4px;background:linear-gradient(90deg,var(--blue-lt),var(--blue));transition:width .4s}}
    .two-col{display:grid;grid-template-columns:1fr 1fr;gap:20px}}
    .overlay{position:fixed;inset:0;background:rgba(20,30,70,.45);display:flex;align-items:center;justify-content:center;z-index:200;animation:fadeIn .15s ease;backdrop-filter:blur(2px)}}
    @keyframes fadeIn{from{opacity:0}}to{opacity:1}}
    .modal{background:#fff;border:1px solid var(--border);border-radius:14px;padding:28px;width:520px;max-height:92vh;overflow-y:auto;animation:slideUp .2s ease;box-shadow:var(--shadow-md)}}
    @keyframes slideUp{from{transform:translateY(16px);opacity:0}}to{transform:translateY(0);opacity:1}}
    .modal-title{font-size:1.1rem;font-weight:700;color:var(--navy);margin-bottom:16px}}
    .modal-actions{display:flex;gap:10px;justify-content:flex-end;margin-top:20px;padding-top:14px;border-top:1px solid var(--border-lt)}}
    .kv-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:16px}}
    .kv{display:flex;flex-direction:column;gap:3px;background:#f7f9fd;border-radius:8px;padding:10px 12px}}
    .kv-key{font-size:.69rem;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}}
    .kv-val{font-size:.88rem;font-weight:600;color:var(--navy)}}
    .toast{position:fixed;bottom:24px;right:24px;background:var(--navy);color:#fff;padding:12px 20px;border-radius:10px;font-size:.84rem;font-weight:600;z-index:999;animation:slideUp .2s ease;display:flex;align-items:center;gap:8px}}
    .saving-row td{opacity:.5}}
    .atraso-chip{display:inline-block;padding:2px 9px;border-radius:6px;font-size:.7rem;font-weight:600}}
  </style>
</head>
<body>
<div id="root"></div>
<script type="text/babel">
const { useState, useEffect, useCallback, useMemo }} = React;
const TODAY = new Date().toISOString().split('T')[0];

/* ── DATOS ── */
const FACTURAS_RAW = __FJ__;

/* ── HELPERS ── */
const fmtM = n => {
  if (!n) return '$0';
  if (n >= 1e9) return '$' + (n/1e9).toFixed(2) + 'B';
  if (n >= 1e6) return '$' + (n/1e6).toFixed(1) + 'M';
  if (n >= 1e3) return '$' + Math.round(n/1e3).toLocaleString('es-CL') + 'K';
  return '$' + Math.round(n).toLocaleString('es-CL');
}};
const fmtFull = n => '$' + Math.round(n||0).toLocaleString('es-CL');
const fmtDate = d => {
  if (!d) return '—';
  try { return new Date(d + 'T12:00:00').toLocaleDateString('es-CL', {day:'2-digit',month:'short',year:'numeric'}}); }} catch(e) { return d; }}
}};
const diasAtraso = fecha => {
  if (!fecha) return 0;
  return Math.floor((new Date(TODAY) - new Date(fecha + 'T12:00:00')) / 86400000);
}};
const estadoNorm = e => (e||'').toLowerCase().trim();

/* ── BADGE ESTADO ── */
function EstadoBadge({estado}}) {
  const e = estadoNorm(estado);
  const map = {
    'pagada':  ['Pagada',  'badge-success'],
    'pagado':  ['Pagada',  'badge-success'],
    'impaga':  ['Impaga',  'badge-warn'],
    'pendiente':['Pendiente','badge-warn'],
    'vencida': ['Vencida', 'badge-danger'],
    'anulada': ['Anulada', 'badge-muted'],
  }};
  const [label, cls] = map[e] || [estado || '—', 'badge-muted'];
  return <span className={{`badge ${cls}}`}}>{label}}</span>;
}}

function SiiBadge({sii}}) {
  const v = (sii||'').toLowerCase().trim();
  if (!v || v === '—') return null;
  const ok = v === 'sí' || v === 'si' || v === 'yes' || v === '1' || v === 'true';
  return <span className={{`badge ${ok ? 'badge-success' : 'badge-danger'}}`}}>{ok ? '✓ SII' : '✗ SII'}}</span>;
}}

function AtrasoChip({fecha, estado}}) {
  const e = estadoNorm(estado);
  if (e === 'pagada' || e === 'pagado' || e === 'anulada') return null;
  const d = diasAtraso(fecha);
  if (d <= 0) return <span className="atraso-chip" style={{background:'#edfaf4',color:'#1a9e6a'}}>Vigente</span>;
  if (d <= 30) return <span className="atraso-chip" style={{background:'#fff7ed',color:'#d97b10'}}>{d}}d atraso</span>;
  return <span className="atraso-chip" style={{background:'#fff0f0',color:'#d93b3b'}}>{d}}d atraso</span>;
}}

/* ── LOGO ── */
function LogoIcon() {
  return (
    <svg viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M8 14C8 11 10 9 13 9H20L14 17H8V14Z" fill="#5fb3f0"/>
      <path d="M8 17H14L8 25V17Z" fill="#4a9de0"/>
      <path d="M14 17L20 9H27C30 9 32 11 32 14V17L26 25H14L20 17H26V14H20L14 17Z" fill="#3a6fd8"/>
      <path d="M14 25H26L20 33H13C10 33 8 31 8 28V25L14 25Z" fill="#5fb3f0"/>
      <path d="M26 25L32 17V28C32 31 30 33 27 33H20L26 25Z" fill="#3a6fd8"/>
    </svg>
  );
}}

/* ── TOAST ── */
function Toast({msg, onDone}}) {
  useEffect(() => { const t = setTimeout(onDone, 3000); return () => clearTimeout(t); }}, []);
  return <div className="toast">{msg}}</div>;
}}

/* ── MODAL CONFIRMAR PAGO ── */
function ModalPago({factura, onClose, onConfirm, saving}}) {
  return (
    <div className="overlay" onClick={{e => e.target===e.currentTarget && onClose()}}>
      <div className="modal">
        <p className="modal-title">💰 Confirmar Pago</p>
        <div className="kv-grid">
          <div className="kv"><span className="kv-key">Cliente</span><span className="kv-val">{factura.razon}}</span></div>
          <div className="kv"><span className="kv-key">N° Factura</span><span className="kv-val">{factura.nFactura}}</span></div>
          <div className="kv"><span className="kv-key">Total</span><span className="kv-val" style={{color:'var(--success)'}}>{fmtFull(factura.total)}</span></div>
          <div className="kv"><span className="kv-key">Fecha Factura</span><span className="kv-val">{fmtDate(factura.fecha)}}</span></div>
        </div>
        <p style={{fontSize:'.84rem',color:'var(--text-2)',lineHeight:1.6}}>
          Al confirmar, se actualizará el estado a <strong>Pagada</strong> directamente en el Excel de SharePoint (<strong>Facturación EST {factura.year}}.xlsx</strong>, fila {factura.row}}).
        </p>
        <div className="modal-actions">
          <button className="btn btn-ghost" onClick={{onClose}} disabled={{saving}}>Cancelar</button>
          <button className="btn btn-success" onClick={{onConfirm}} disabled={{saving}}>
            {saving ? '⏳ Guardando...' : '✓ Confirmar Pago'}}
          </button>
        </div>
      </div>
    </div>
  );
}}

/* ── API MARCAR PAGADA ── */
async function marcarPagadaAPI(factura) {
  const resp = await fetch('/api/pagar', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' }},
    body: JSON.stringify({
      year:     factura.year,
      row:      factura.row,
      nFactura: factura.nFactura
    }})
  }});
  if (!resp.ok) throw new Error(await resp.text());
  return resp.json();
}}

/* ══════════════════════════════════════════
   DASHBOARD
══════════════════════════════════════════ */
function Dashboard({facturas, setView}}) {
  const activas = facturas.filter(f => estadoNorm(f.estado) !== 'anulada');
  const pagadas = activas.filter(f => ['pagada','pagado'].includes(estadoNorm(f.estado)));
  const impagas = activas.filter(f => !['pagada','pagado','anulada'].includes(estadoNorm(f.estado)));
  const vencidas = impagas.filter(f => f.fecha && diasAtraso(f.fecha) > (parseInt(f.plazo)||30));

  const totalFacturado = activas.reduce((s,f) => s + (f.total||0), 0);
  const totalPagado    = pagadas.reduce((s,f) => s + (f.total||0), 0);
  const totalPendiente = impagas.reduce((s,f) => s + (f.total||0), 0);
  const recup = totalFacturado > 0 ? Math.round(totalPagado/totalFacturado*100) : 0;

  // Top clientes por pendiente
  const porCliente = {}};
  impagas.forEach(f => {
    if (!porCliente[f.razon]) porCliente[f.razon] = {razon: f.razon, rut: f.rut, pendiente: 0, nImpagas: 0}};
    porCliente[f.razon].pendiente += f.total||0;
    porCliente[f.razon].nImpagas++;
  }});
  const top5 = Object.values(porCliente).sort((a,b) => b.pendiente - a.pendiente).slice(0,5);

  return (
    <>
      <div className="page-header">
        <p className="page-title">Dashboard</p>
        <p className="page-sub">Trayectoria EST · {activas.length}} facturas activas · Actualizado __TODAY__</p>
      </div>
      <div className="stats-grid">
        {[
          {label:'Total Facturado', val:fmtM(totalFacturado), sub:fmtFull(totalFacturado), icon:'📋', bg:'#eef4fd', ic:'var(--blue)'}},
          {label:'Por Cobrar',      val:fmtM(totalPendiente), sub:fmtFull(totalPendiente), icon:'⏳', bg:'#fff7ed', ic:'var(--warn)'}},
          {label:'Recuperado',      val:fmtM(totalPagado),   sub:`${recup}}% del total`,  icon:'✅', bg:'#edfaf4', ic:'var(--success)'}},
          {label:'Vencidas',        val:vencidas.length,      sub:fmtM(vencidas.reduce((s,f)=>s+(f.total||0),0))+' en riesgo', icon:'⚠️', bg:'#fff0f0', ic:'var(--danger)'}},
        ].map(c => (
          <div key={{c.label}} className="stat-card">
            <div className="stat-card-top">
              <div className="stat-label">{c.label}}</div>
              <div className="stat-icon" style={{background:c.bg}}>{c.icon}}</div>
            </div>
            <div className="stat-val">{c.val}}</div>
            <div className="stat-sub">{c.sub}}</div>
          </div>
        ))}}
      </div>
      <div className="two-col">
        <div className="table-wrap">
          <div className="table-header">
            <span className="table-title">Top 5 Clientes — Monto Pendiente</span>
          </div>
          <table>
            <thead><tr><th>#</th><th>Cliente</th><th>Pendiente</th><th>Facturas</th></tr></thead>
            <tbody>{top5.map((c,i) => (
              <tr key={{c.razon}}>
                <td style={{fontWeight:700,color:'var(--muted)'}}>#{i+1}}</td>
                <td><div style={{fontWeight:600,color:'var(--navy)'}}> {c.razon}}</div><div style={{fontSize:'.74rem',color:'var(--muted)'}}> {c.rut}}</div></td>
                <td style={{fontWeight:700,color:'var(--danger)'}}> {fmtM(c.pendiente)}}</td>
                <td><span style={{fontWeight:700,color:'var(--warn)'}}> {c.nImpagas}}</span> impaga(s)</td>
              </tr>
            ))}}</tbody>
          </table>
        </div>
        <div className="table-wrap">
          <div className="table-header">
            <span className="table-title">Facturas Vencidas Recientes</span>
            <button className="btn btn-ghost btn-sm" onClick={{()=>setView('facturas')}}>Ver todas →</button>
          </div>
          <table>
            <thead><tr><th>N° Factura</th><th>Cliente</th><th>Total</th><th>Atraso</th></tr></thead>
            <tbody>
              {vencidas.slice(0,6).map(f => (
                <tr key={{f.nFactura+f.row}}>
                  <td style={{fontWeight:700}}> {f.nFactura}}</td>
                  <td style={{fontSize:'.8rem'}}>{f.razon}}</td>
                  <td style={{fontWeight:700,color:'var(--danger)'}}> {fmtM(f.total)}}</td>
                  <td><AtrasoChip fecha={{f.fecha}} estado={{f.estado}}/></td>
                </tr>
              ))}}
              {vencidas.length===0 && <tr><td colSpan="4"><div className="empty-state" style={{padding:20}}>Sin facturas vencidas 🎉</div></td></tr>}}
            </tbody>
          </table>
        </div>
      </div>
    </>
  );
}}

/* ══════════════════════════════════════════
   FACTURAS
══════════════════════════════════════════ */
function FacturasView({facturas, onPagar}}) {
  const [search,   setSearch]   = useState('');
  const [efEstado, setEfEstado] = useState('todas');
  const [efYear,   setEfYear]   = useState('todos');

  const years  = [...new Set(facturas.map(f => f.year))].sort();
  const estados = ['todas','impaga','pagada','vencida','anulada'];

  const fil = useMemo(() => {
    return facturas.filter(f => {
      const txt = (f.razon + f.rut + f.nFactura + f.oc).toLowerCase();
      if (!txt.includes(search.toLowerCase())) return false;
      if (efYear !== 'todos' && f.year !== parseInt(efYear)) return false;
      if (efEstado === 'todas') return true;
      const e = estadoNorm(f.estado);
      if (efEstado === 'impaga')  return !['pagada','pagado','anulada'].includes(e);
      if (efEstado === 'pagada')  return ['pagada','pagado'].includes(e);
      if (efEstado === 'anulada') return e === 'anulada';
      if (efEstado === 'vencida') return !['pagada','pagado','anulada'].includes(e) && diasAtraso(f.fecha) > (parseInt(f.plazo)||30);
      return true;
    }}).sort((a,b) => (b.fecha||'').localeCompare(a.fecha||''));
  }}, [facturas, search, efEstado, efYear]);

  const cnt = {
    todas:   facturas.length,
    impaga:  facturas.filter(f=>!['pagada','pagado','anulada'].includes(estadoNorm(f.estado))).length,
    pagada:  facturas.filter(f=>['pagada','pagado'].includes(estadoNorm(f.estado))).length,
    vencida: facturas.filter(f=>!['pagada','pagado','anulada'].includes(estadoNorm(f.estado))&&diasAtraso(f.fecha)>(parseInt(f.plazo)||30)).length,
    anulada: facturas.filter(f=>estadoNorm(f.estado)==='anulada').length,
  }};

  return (
    <>
      <div className="page-header">
        <p className="page-title">Facturas</p>
        <p className="page-sub">Todas las facturas desde SharePoint</p>
      </div>
      <div className="filter-bar">
        {estados.map(e => (
          <button key={{e}} className={{`btn btn-sm ${efEstado===e?'btn-primary':'btn-ghost'}}`}} onClick={{()=>setEfEstado(e)}}>
            {e.charAt(0).toUpperCase()+e.slice(1)}} ({cnt[e]||0}})
          </button>
        ))}}
        <select className="select-bar" value={{efYear}} onChange={{e=>setEfYear(e.target.value)}}>
          <option value="todos">Todos los años</option>
          {years.map(y=><option key={{y}} value={{y}}>{y}}</option>)}}
        </select>
      </div>
      <div className="table-wrap">
        <div className="table-header">
          <span className="table-title">{fil.length}} facturas</span>
          <input className="search-bar" placeholder="🔍 Buscar cliente, RUT, N° factura..." value={{search}} onChange={{e=>setSearch(e.target.value)}}/>
        </div>
        <table>
          <thead>
            <tr>
              <th>N° Factura</th><th>Cliente</th><th>RUT</th>
              <th>Fecha</th><th>Neto</th><th>Total c/IVA</th>
              <th>Estado</th><th>Plazo</th><th>Fecha Pago</th>
              <th>OC</th><th>SII</th><th></th>
            </tr>
          </thead>
          <tbody>
            {fil.map(f => {
              const esImpaga = !['pagada','pagado','anulada'].includes(estadoNorm(f.estado));
              const vencida  = esImpaga && diasAtraso(f.fecha) > (parseInt(f.plazo)||30);
              return (
                <tr key={{f.nFactura+'-'+f.row+'-'+f.year}} style={{background: vencida?'#fff8f8':esImpaga?'#fffdf5':''}}>
                  <td style={{fontWeight:700,color:'var(--navy)'}}> {f.nFactura}}</td>
                  <td><div style={{fontWeight:600,fontSize:'.82rem'}}> {f.razon}}</div></td>
                  <td style={{color:'var(--muted)',fontSize:'.78rem'}}> {f.rut}}</td>
                  <td style={{fontSize:'.78rem',color:'var(--text-2)'}}> {fmtDate(f.fecha)}}</td>
                  <td style={{fontSize:'.82rem'}}> {fmtFull(f.neto)}}</td>
                  <td style={{fontWeight:700,color:vencida?'var(--danger)':esImpaga?'var(--warn)':'var(--success)'}}> {fmtFull(f.total)}}</td>
                  <td><EstadoBadge estado={{f.estado}}/></td>
                  <td style={{fontSize:'.78rem',color:'var(--muted)'}}> {f.plazo ? f.plazo+' días' : '—'}}</td>
                  <td style={{fontSize:'.78rem'}}> {fmtDate(f.fechaPago)}}</td>
                  <td style={{fontSize:'.76rem',color:'var(--muted)'}}> {f.oc||'—'}}</td>
                  <td><SiiBadge sii={{f.sii}}/></td>
                  <td>
                    {esImpaga && (
                      <button className="btn btn-success btn-sm" title="Marcar como pagada" onClick={{()=>onPagar(f)}}>
                        ✓ Pagar
                      </button>
                    )}}
                  </td>
                </tr>
              );
            }})}}
            {fil.length===0 && <tr><td colSpan="12"><div className="empty-state">Sin facturas que mostrar.</div></td></tr>}}
          </tbody>
        </table>
      </div>
    </>
  );
}}

/* ══════════════════════════════════════════
   CLIENTES
══════════════════════════════════════════ */
function ClientesView({facturas}}) {
  const [search, setSearch] = useState('');

  const clientes = useMemo(() => {
    const map = {}};
    facturas.filter(f=>estadoNorm(f.estado)!=='anulada').forEach(f => {
      if (!map[f.rut]) map[f.rut] = {razon:f.razon, rut:f.rut, total:0, pendiente:0, pagado:0, nFacturas:0, nImpagas:0}};
      map[f.rut].nFacturas++;
      map[f.rut].total += f.total||0;
      if (['pagada','pagado'].includes(estadoNorm(f.estado))) {
        map[f.rut].pagado += f.total||0;
      }} else {
        map[f.rut].pendiente += f.total||0;
        map[f.rut].nImpagas++;
      }}
    }});
    return Object.values(map).sort((a,b) => b.pendiente - a.pendiente);
  }}, [facturas]);

  const fil = clientes.filter(c =>
    c.razon.toLowerCase().includes(search.toLowerCase()) || c.rut.includes(search)
  );

  return (
    <>
      <div className="page-header">
        <p className="page-title">Clientes</p>
        <p className="page-sub">{clientes.length}} clientes con facturas activas</p>
      </div>
      <div className="table-wrap">
        <div className="table-header">
          <span className="table-title">{fil.length}} clientes</span>
          <input className="search-bar" placeholder="🔍 Buscar nombre o RUT..." value={{search}} onChange={{e=>setSearch(e.target.value)}}/>
        </div>
        <table>
          <thead><tr><th>Cliente</th><th>RUT</th><th>Facturas</th><th>Total Facturado</th><th>Pendiente</th><th>Recuperación</th></tr></thead>
          <tbody>{fil.map(c => {
            const pct = c.total > 0 ? Math.round(c.pagado/c.total*100) : 0;
            return (
              <tr key={{c.rut}}>
                <td style={{fontWeight:600,color:'var(--navy)'}}> {c.razon}}</td>
                <td style={{color:'var(--muted)',fontSize:'.78rem'}}> {c.rut}}</td>
                <td><span style={{fontWeight:700,color:'var(--warn)'}}> {c.nImpagas}}</span><span style={{color:'var(--muted)',fontSize:'.78rem'}}> imp. / {c.nFacturas}}</span></td>
                <td style={{fontSize:'.82rem'}}> {fmtM(c.total)}}</td>
                <td style={{fontWeight:700,color:'var(--danger)'}}> {fmtM(c.pendiente)}}</td>
                <td style={{width:130}}>
                  <div style={{display:'flex',justifyContent:'space-between',fontSize:'.72rem',color:'var(--muted)',marginBottom:3}}><span>{pct}}%</span></div>
                  <div className="progress-bar"><div className="progress-fill" style={{width:pct+'%'}}/></div>
                </td>
              </tr>
            );
          }})}}
          </tbody>
        </table>
      </div>
    </>
  );
}}

/* ══════════════════════════════════════════
   APP
══════════════════════════════════════════ */
function App() {
  const [facturas, setFacturas] = useState(FACTURAS_RAW);
  const [view,     setView]     = useState('dashboard');
  const [modal,    setModal]    = useState(null);  // factura a pagar
  const [saving,   setSaving]   = useState(false);
  const [toast,    setToast]    = useState(null);

  const impagas = facturas.filter(f => !['pagada','pagado','anulada'].includes(estadoNorm(f.estado)));
  const vencCnt = facturas.filter(f => !['pagada','pagado','anulada'].includes(estadoNorm(f.estado)) && diasAtraso(f.fecha) > (parseInt(f.plazo)||30)).length;

  const handlePagar = (factura) => setModal(factura);

  const handleConfirmarPago = async () => {
    setSaving(true);
    try {
      await marcarPagadaAPI(modal);
      // Actualizar estado local inmediatamente
      setFacturas(prev => prev.map(f =>
        f.row === modal.row && f.year === modal.year
          ? {...f, estado: 'Pagada', fechaPago: TODAY}}
          : f
      ));
      setToast('✅ Factura ' + modal.nFactura + ' marcada como pagada en SharePoint');
      setModal(null);
    }} catch(err) {
      setToast('❌ Error: ' + err.message);
    }} finally {
      setSaving(false);
    }}
  }};

  return (
    <div className="crm-root">
      <nav className="sidebar">
        <div className="sidebar-logo">
          <div className="logo-mark">
            <div className="logo-icon"><LogoIcon/></div>
            <div className="logo-text">Trayectoria EST<span>CRM Facturación</span></div>
          </div>
        </div>
        <div className="nav-section">Menú</div>
        {[
          {id:'dashboard', icon:'📊', label:'Dashboard'}},
          {id:'facturas',  icon:'🧾', label:'Facturas', badge: impagas.length > 0 ? impagas.length : null}},
          {id:'clientes',  icon:'🏢', label:'Clientes'}},
        ].map(n => (
          <button key={{n.id}} className={{`nav-btn ${view===n.id?'active':''}}`}} onClick={{()=>setView(n.id)}}>
            <span className="nav-icon">{n.icon}}</span>
            {n.label}}
            {n.badge && <span className="nav-badge">{n.badge}}</span>}}
          </button>
        ))}}
        <div className="sync-info">
          🔄 Sincronizado desde SharePoint<br/>
          Actualizado __TODAY__ a las 08:00
        </div>
      </nav>

      <main className="main">
        {view === 'dashboard' && <Dashboard facturas={{facturas}} setView={{setView}}/>}}
        {view === 'facturas'  && <FacturasView facturas={{facturas}} onPagar={{handlePagar}}/>}}
        {view === 'clientes'  && <ClientesView facturas={{facturas}}/>}}
      </main>

      {modal && (
        <ModalPago
          factura={{modal}}
          onClose={{()=>setModal(null)}}
          onConfirm={{handleConfirmarPago}}
          saving={{saving}}
        />
      )}}

      {toast && <Toast msg={{toast}} onDone={{()=>setToast(null)}}/>}}
    </div>
  );
}}

ReactDOM.createRoot(document.getElementById('root')).render(<App/>);
</script>
</body>
</html>"""
    html = html.replace('__FJ__', fj).replace('__TODAY__', today_str)
    return html

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    print("🔐 Obteniendo token Microsoft Graph...")
    token = get_token()

    all_facturas = []
    for year, filename in FILES.items():
        print(f"📥 Leyendo {filename}...")
        data = get_file_bytes(token, year, filename)
        if data:
            rows = parse_excel(data, year)
            all_facturas.extend(rows)
            print(f"   ✅ {len(rows)} filas leídas")
        else:
            print(f"   ⚠️  Skipping {filename}")

    print(f"\n📊 Total: {len(all_facturas)} facturas")
    print("🏗️  Generando index.html...")
    html = build_html(all_facturas)
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ index.html generado correctamente")

if __name__ == "__main__":
    main()
