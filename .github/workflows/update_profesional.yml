#!/usr/bin/env python3
"""
update_profesional.py
Descarga archivos de nómina Profesional desde Google Drive,
recalcula proyecciones y actualiza los datos embebidos en el HTML.
"""

import json, re, sys, requests, xlrd, unicodedata
from datetime import datetime
from collections import defaultdict

def normalize(s):
    """Elimina acentos y pasa a minúsculas para comparaciones robustas."""
    return unicodedata.normalize('NFD', s).encode('ascii', 'ignore').decode('ascii').lower()

# ── Configuración ─────────────────────────────────────────────────
API_KEY   = "AIzaSyAId7gthv7EEzmaTrfbt07FK4Kf-ii51uM"
FOLDER_ID = "1QWRO2A3eO4Aa5x95IVtIgGV_AkD1I_Uj"
HTML_FILE = "index.html"

METAS = {
    1: 3156777,  2: 4481604,  3: 4842093,  4: 5323095,
    5: 5324319,  6: 5471620,  7: 5415600,  8: 5315794,
    9: 5577614, 10: 5656792, 11: 7134201, 12: 5783355
}

MONTH_NAMES = {
    "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
    "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12
}

MONTH_LABELS = {
    1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
    7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"
}

# ── Drive helpers ─────────────────────────────────────────────────
def drive_list_files(folder_id):
    resp = requests.get(
        "https://www.googleapis.com/drive/v3/files",
        params={
            "q": f"'{folder_id}' in parents and trashed=false",
            "fields": "files(id,name,mimeType)",
            "pageSize": 100,
            "key": API_KEY
        },
        timeout=30
    )
    resp.raise_for_status()
    return resp.json().get("files", [])

def drive_download(file_id):
    resp = requests.get(
        f"https://www.googleapis.com/drive/v3/files/{file_id}",
        params={"alt": "media", "key": API_KEY},
        timeout=60
    )
    resp.raise_for_status()
    return resp.content

# ── Columnas fijas (base 0) ───────────────────────────────────────
# Archivo NÓMINA
NOM_RFC     = 0   # Columna A
NOM_CONTRIB = 1   # Columna B
NOM_PERIODO = 5   # Columna F
NOM_M       = 12  # Columna M (resta)
NOM_O       = 14  # Columna O (resta)
NOM_R       = 17  # Columna R (suma)
# Recaudación nómina = R − O − M

# Archivo RETENCIÓN
RET_RFC     = 0   # Columna A
RET_CONTRIB = 1   # Columna B
RET_PERIODO = 4   # Columna E
RET_IMP     = 13  # Columna N (IMPUESTO)

def _find_data_start(ws, rfc_col):
    for i in range(min(20, ws.nrows)):
        val = str(ws.cell_value(i, rfc_col)).strip()
        if len(val) >= 12 and not val.replace(" ", "").isalpha():
            return i
    return -1

def _parse_nomina(ws):
    data_start = _find_data_start(ws, NOM_RFC)
    if data_start < 0:
        print("  No se encontró fila de datos (nómina)", file=sys.stderr)
        return []
    records = []
    for i in range(data_start, ws.nrows):
        try:
            rfc = str(ws.cell_value(i, NOM_RFC)).strip().upper()
            if not rfc or len(rfc) < 12:
                continue
            periodo_raw = ws.cell_value(i, NOM_PERIODO)
            try:
                periodo = str(int(float(str(periodo_raw))))
            except Exception:
                periodo = str(periodo_raw).strip()
            if len(periodo) != 6:
                continue
            val_r   = float(ws.cell_value(i, NOM_R)) if ws.ncols > NOM_R else 0.0
            val_o   = float(ws.cell_value(i, NOM_O)) if ws.ncols > NOM_O else 0.0
            val_m   = float(ws.cell_value(i, NOM_M)) if ws.ncols > NOM_M else 0.0
            contrib = str(ws.cell_value(i, NOM_CONTRIB)).strip() if ws.ncols > NOM_CONTRIB else ""
            records.append({"rfc": rfc, "periodo": periodo,
                             "recaudacion": val_r - val_o - val_m, "contrib": contrib})
        except Exception:
            continue
    return records

def _parse_retencion(ws):
    data_start = _find_data_start(ws, RET_RFC)
    if data_start < 0:
        print("  No se encontró fila de datos (retención)", file=sys.stderr)
        return []
    records = []
    for i in range(data_start, ws.nrows):
        try:
            rfc = str(ws.cell_value(i, RET_RFC)).strip().upper()
            if not rfc or len(rfc) < 12:
                continue
            periodo_raw = ws.cell_value(i, RET_PERIODO)
            try:
                periodo = str(int(float(str(periodo_raw))))
            except Exception:
                periodo = str(periodo_raw).strip()
            if len(periodo) != 6:
                continue
            imp     = float(ws.cell_value(i, RET_IMP)) if ws.ncols > RET_IMP else 0.0
            contrib = str(ws.cell_value(i, RET_CONTRIB)).strip() if ws.ncols > RET_CONTRIB else ""
            records.append({"rfc": rfc, "periodo": periodo,
                             "recaudacion": imp, "contrib": contrib})
        except Exception:
            continue
    return records

def parse_xls(file_bytes, file_type="nomina"):
    try:
        wb = xlrd.open_workbook(file_contents=file_bytes)
    except Exception as e:
        print(f"  xlrd error: {e}", file=sys.stderr)
        return []
    ws = wb.sheet_by_index(0)
    if file_type == "retencion":
        return _parse_retencion(ws)
    return _parse_nomina(ws)

# ── Lógica de omisos ──────────────────────────────────────────────
def prev_period(p):
    p = str(p)
    year, month = int(p[:4]), int(p[4:])
    month -= 1
    if month == 0:
        month, year = 12, year - 1
    return f"{year}{month:02d}"

def format_period(p):
    s = str(p)
    meses = ['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
    try:
        return meses[int(s[4:6])] + '-' + s[2:4]
    except Exception:
        return s

def get_dominant(records):
    totals = defaultdict(float)
    for r in records:
        if r['periodo']:
            totals[r['periodo']] += r['recaudacion']
    return max(totals, key=totals.get) if totals else None

def get_missing_periods(paid_set, dominant, max_back=12, stop_before=None):
    missing, p = [], str(dominant)
    while p not in paid_set:
        if stop_before and int(p) < int(stop_before):
            break
        missing.append(p)
        p = prev_period(p)
        if len(missing) >= max_back:
            break
    return missing

def compute_month(month_num, all_month_data):
    cur      = all_month_data.get(month_num, [])
    dominant = get_dominant(cur)
    acumulado = sum(r['recaudacion'] for r in cur)

    # Meses de referencia: hasta 4 anteriores con datos
    ref_months = [m for m in range(max(1, month_num - 4), month_num)
                  if all_month_data.get(m)]
    n_ref = len(ref_months)

    # Periodos globales (todos los meses hasta este)
    global_periods = defaultdict(dict)
    global_contrib = {}
    for m in range(1, month_num + 1):
        for r in all_month_data.get(m, []):
            gp = global_periods[r['rfc']]
            if r['periodo'] not in gp or r['recaudacion'] > gp[r['periodo']]:
                gp[r['periodo']] = r['recaudacion']
            if r['rfc'] not in global_contrib and r['contrib']:
                global_contrib[r['rfc']] = r['contrib']

    # Conteo de meses de referencia y montos por RFC
    rfc_ref_count   = defaultdict(int)
    rfc_ref_periods = defaultdict(dict)
    rfc_contrib     = {}
    paid_2026_in_ref = defaultdict(set)

    for rm in ref_months:
        seen = set()
        for r in all_month_data.get(rm, []):
            rp = rfc_ref_periods[r['rfc']]
            if r['periodo'] not in rp or r['recaudacion'] > rp[r['periodo']]:
                rp[r['periodo']] = r['recaudacion']
            if str(r['periodo']).startswith('2026'):
                paid_2026_in_ref[r['rfc']].add(str(r['periodo']))
            if r['rfc'] not in seen:
                seen.add(r['rfc'])
                rfc_ref_count[r['rfc']] += 1
            if r['rfc'] not in rfc_contrib and r['contrib']:
                rfc_contrib[r['rfc']] = r['contrib']

    # Candidatos: ≥2 meses ref O ≥2 periodos 2026 en meses ref
    candidates = {rfc for rfc, cnt in rfc_ref_count.items() if cnt >= 2}
    candidates |= {rfc for rfc, ps in paid_2026_in_ref.items() if len(ps) >= 2}

    omisos = []
    for rfc in candidates:
        cnt          = rfc_ref_count.get(rfc, 0)
        periods_2026 = paid_2026_in_ref.get(rfc, set())
        if cnt < 2 and len(periods_2026) < 2:
            continue

        paid_set = set(global_periods[rfc].keys())
        if not dominant or dominant in paid_set:
            continue

        has_2026 = any(p.startswith('2026') for p in global_periods[rfc])
        missing  = (get_missing_periods(paid_set, dominant, 12, '202601')
                    if not has_2026
                    else get_missing_periods(paid_set, dominant, 12))
        if not missing:
            continue

        ref_amounts = list(rfc_ref_periods[rfc].values())
        if not ref_amounts:
            continue

        avg = sum(ref_amounts) / len(ref_amounts)
        est = avg * len(missing)

        seg = ('omisos_totales' if not has_2026
               else 'alta'       if cnt >= n_ref
               else 'media'      if cnt >= 3
               else 'seguimiento')

        contrib        = rfc_contrib.get(rfc) or global_contrib.get(rfc) or ''
        pending_labels = [format_period(p) for p in missing]

        omisos.append({
            'rfc': rfc, 'contrib': contrib, 'count': cnt,
            'avg': round(est), 'n_missing': len(missing),
            'pending': pending_labels, 'seg': seg
        })

    omisos.sort(key=lambda o: -o['avg'])
    esperado    = sum(o['avg'] for o in omisos if o['seg'] in ('alta', 'media'))
    proyeccion  = acumulado + esperado
    meta        = METAS.get(month_num, 0)

    # Segmentos
    segments = {}
    for o in omisos:
        key = o['seg']
        if key not in segments:
            segments[key] = {'count': 0, 'monto': 0, 'omisos': []}
        segments[key]['count'] += 1
        segments[key]['monto'] += o['avg']
        segments[key]['omisos'].append({
            'rfc': o['rfc'], 'contrib': o['contrib'], 'avg': o['avg'],
            'count': o['count'], 'nMissing': o['n_missing'], 'pending': o['pending']
        })
    for s in segments.values():
        s['monto'] = round(s['monto'])
        s['omisos'].sort(key=lambda x: -x['avg'])

    return {
        'mes_label':         MONTH_LABELS.get(month_num, str(month_num)),
        'mes_num':           month_num,
        'meta':              meta,
        'dominant_period':   int(dominant) if dominant else 0,
        'ref_months':        ref_months,
        'acumulado_real':    round(acumulado),
        'total_omisos':      len(omisos),
        'total_esperado':    round(esperado),
        'proyeccion_cierre': round(proyeccion),
        'meta_cruzada':      proyeccion >= meta,
        'pct_acumulado':     acumulado / meta * 100 if meta else 0,
        'pct_proyeccion':    proyeccion / meta * 100 if meta else 0,
        'segmentos':         segments,
        'omisos': [
            {'rfc': o['rfc'], 'contrib': o['contrib'], 'avg': o['avg'],
             'count': o['count'], 'nMissing': o['n_missing'],
             'pending': o['pending'], 'seg': o['seg']}
            for o in omisos[:5000]
        ]
    }

# ── Main ──────────────────────────────────────────────────────────
def main():
    print("Listando archivos en Drive (Profesional)...")
    raw_files = drive_list_files(FOLDER_ID)

    files = []
    for f in raw_files:
        name     = (f.get('name') or '').lower()
        name_n   = normalize(f.get('name') or '')   # sin acentos
        mime     = f.get('mimeType', '')
        if mime == 'application/vnd.ms-excel' or name.endswith('.xls'):
            month_name = next((m for m in MONTH_NAMES if m in name_n), None)
            if month_name:
                file_type = 'retencion' if 'retencion' in name_n else 'nomina'
                print(f"  Detectado: {f['name']} → mes={month_name}, tipo={file_type}")
                files.append({'id': f['id'], 'name': f['name'],
                               'num': MONTH_NAMES[month_name], 'type': file_type})

    files.sort(key=lambda f: (f['num'], f['type']))

    if not files:
        print("ERROR: No se encontraron archivos .xls en la carpeta", file=sys.stderr)
        sys.exit(1)

    print(f"Encontrados {len(files)} archivos: {[f['name'] for f in files]}")

    # Descargar y parsear — combinar nómina + retención por mes
    all_month_data = defaultdict(list)
    for fi in files:
        print(f"  Descargando {fi['name']} [{fi['type']}]...")
        try:
            raw     = drive_download(fi['id'])
            records = parse_xls(raw, fi['type'])
            print(f"    {len(records)} registros")
            all_month_data[fi['num']].extend(records)
        except Exception as e:
            print(f"    ERROR: {e}", file=sys.stderr)

    # Calcular proyecciones
    print("Calculando proyecciones...")
    all_data   = {}
    month_nums = sorted(all_month_data.keys())
    for num in month_nums:
        print(f"  {MONTH_LABELS[num]}...", end=" ")
        d = compute_month(num, all_month_data)
        all_data[str(num)] = d
        print(f"acumulado=${d['acumulado_real']:,.0f}  omisos={d['total_omisos']}  est=${d['total_esperado']:,.0f}")

    # Validar datos
    if not any(d.get('acumulado_real', 0) > 0 for d in all_data.values()):
        print("ERROR: todos los meses tienen acumulado=0 — abortando", file=sys.stderr)
        sys.exit(1)

    # Actualizar HTML
    print(f"\nActualizando {HTML_FILE}...")
    with open(HTML_FILE, 'r', encoding='utf-8') as f:
        html = f.read()

    new_json = json.dumps(all_data, ensure_ascii=False, separators=(',', ':'))

    # Reemplazar allData línea por línea
    lines     = html.split('\n')
    new_lines = []
    n1        = 0
    for line in lines:
        if re.match(r'\s*(var|let|const)\s+allData\s*=\s*\{', line):
            new_lines.append(f'let allData = {new_json};')
            n1 += 1
        else:
            new_lines.append(line)
    html = '\n'.join(new_lines)

    if n1 == 0:
        print("ERROR: no se encontró 'let allData' en el HTML", file=sys.stderr)
        sys.exit(1)

    # Actualizar lastUpdated
    now = datetime.now().strftime('%-d/%-m/%Y, %H:%M:%S')
    html, _ = re.subn(
        r"(var lastUpdated\s*=\s*')[^']*(')",
        lambda m: m.group(1) + now + m.group(2),
        html
    )

    with open(HTML_FILE, 'w', encoding='utf-8') as f:
        f.write(html)

    latest = max(month_nums)
    d      = all_data[str(latest)]
    print(f"✓ Dashboard Profesional actualizado — {d['mes_label']}: "
          f"${d['acumulado_real']:,.0f} acumulado, "
          f"{d['total_omisos']} omisos, "
          f"${d['proyeccion_cierre']:,.0f} proyección ({d['pct_proyeccion']:.1f}% meta)")

if __name__ == '__main__':
    main()
