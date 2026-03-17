import pandas as pd
import numpy as np
import os
import re
import sys
import time
import traceback
import win32com.client
import tkinter as tk
from tkinter import messagebox
from typing import Optional
try:
from PIL import Image, ImageTk
PIL_DISPONIBLE = True
except ImportError:
PIL_DISPONIBLE = False

# =========================================================

# Braskem Automator v6.7 FINAL

# =========================================================

_USUARIO = os.getenv(“USERNAME”, “usuario”)
_SHAREPOINT = os.path.join(
“C:\Users”, _USUARIO,
“OneDrive - Braskem S.A”,
“Bienvenido al SharePoint de Finanzas México - 3. Tesorería”,
“Reportes Python”
)
CATALOGOS_PATH = os.path.join(_SHAREPOINT, “Catalogos”)
REPORTES_PATH  = os.path.join(_SHAREPOINT, “Reportes”)
os.makedirs(CATALOGOS_PATH, exist_ok=True)
os.makedirs(REPORTES_PATH,  exist_ok=True)

if getattr(sys, “frozen”, False):
_BASE = os.path.dirname(sys.executable)
else:
_BASE = os.path.dirname(os.path.abspath(**file**))

if not os.path.exists(_SHAREPOINT):
CATALOGOS_PATH = os.path.join(_BASE, “Braskem Reporte de Pagos”)
REPORTES_PATH  = os.path.join(_BASE, “Braskem Reporte de Pagos”)
os.makedirs(CATALOGOS_PATH, exist_ok=True)

def resource_path(relative_path: str) -> str:
try:
base_path = getattr(sys, “_MEIPASS”, os.path.dirname(os.path.abspath(**file**)))
except Exception:
base_path = os.path.dirname(os.path.abspath(**file**))
return os.path.join(base_path, relative_path)

FILE_TC       = os.path.join(REPORTES_PATH,  “TC.txt”)
FILE_RP       = os.path.join(REPORTES_PATH,  “RP.txt”)
CATALOGO_PATH = os.path.join(CATALOGOS_PATH, “Catalogo_tesoreria.xlsx”)
OUTPUT_PATH   = os.path.join(REPORTES_PATH,  “Reporte Base.xlsx”)
RESUMEN_PATH  = os.path.join(REPORTES_PATH,  “Reporte de Pagos.xlsx”)

import platform
EJECUTAR_SAP  = True
ZFILL_LEN     = 10
DOC_COL_CANON = “N° documento”

def esperar_archivo(path, timeout=60, min_size=200):
t0 = time.time()
last_size = -1
stable_count = 0
while time.time() - t0 < timeout:
if os.path.exists(path):
try:
size = os.path.getsize(path)
except Exception:
size = 0
if size >= min_size and size == last_size:
stable_count += 1
if stable_count >= 3:
return True
else:
stable_count = 0
last_size = size
time.sleep(1)
return False

def limpiar_monto_sap(v):
if v is None:
return 0.0
s = str(v).strip()
if s == “” or s.lower() == “nan”:
return 0.0
neg = s.endswith(”-”)
if neg:
s = s[:-1]
s = s.strip()
tiene_punto = “.” in s
tiene_coma  = “,” in s
if tiene_punto and tiene_coma:
if s.rindex(”.”) > s.rindex(”,”):
s = s.replace(”,”, “”)
else:
s = s.replace(”.”, “”).replace(”,”, “.”)
elif tiene_coma:
s = s.replace(”,”, “.”)
try:
return -float(s) if neg else float(s)
except Exception:
return 0.0

def normalizar_col(s: str) -> str:
s = str(s).strip().lower()
s = (s.replace(“ó”, “o”).replace(“á”, “a”).replace(“é”, “e”)
.replace(“í”, “i”).replace(“ú”, “u”))
s = re.sub(r”\s+”, “ “, s).strip()
return s

def norm_key(x, zfill_len=ZFILL_LEN):
s = “” if pd.isna(x) else str(x)
s = s.strip()
s = re.sub(r”.0$”, “”, s)
s = re.sub(r”\s+”, “”, s)
m = re.search(r”\d+”, s)
if m:
s = m.group()
if zfill_len:
s = s.zfill(int(zfill_len))
return s

def leer_tabla_sap_pipe(path_txt, encoding=“latin-1”):
if not path_txt or not os.path.exists(path_txt):
return pd.DataFrame()
with open(path_txt, “r”, encoding=encoding, errors=“ignore”) as f:
lines = [ln.rstrip(”\n”) for ln in f]
header_idx = None
for i, ln in enumerate(lines):
if ln.startswith(”|”) and (“Cuenta” in ln or “Mon” in ln or “Tipo cambio” in ln):
header_idx = i
break
if header_idx is None:
return pd.DataFrame()
cols = [c.strip() for c in lines[header_idx].split(”|”)[1:-1]]
data = []
for ln in lines[header_idx + 1:]:
if not ln.startswith(”|”):
continue
core = ln.replace(”|”, “”).strip()
if core and set(core) <= {”-”}:
continue
parts = [p.strip() for p in ln.split(”|”)[1:-1]]
if len(parts) != len(cols):
continue
data.append(parts)
df = pd.DataFrame(data, columns=cols)
df.columns = df.columns.str.strip()
return df

def mostrar_error(titulo, e):
detalle = traceback.format_exc()
messagebox.showerror(titulo, f”{str(e)}\n\n— TRACEBACK —\n{detalle}”)

def popup_exito():
win = tk.Toplevel(root)
win.title(“Proceso completado”)
win.configure(bg=“white”)
win.resizable(False, False)
win.grab_set()
win.geometry(“400x200”)
root.update_idletasks()
x = root.winfo_x() + (root.winfo_width() // 2) - 200
y = root.winfo_y() + (root.winfo_height() // 2) - 100
win.geometry(f”400x200+{x}+{y}”)
tk.Label(win, text=“Proceso completado exitosamente”,
font=(“Arial”, 11, “bold”), fg=”#1A7A1A”, bg=“white”).pack(pady=(20, 6))
tk.Label(win, text=“Archivos guardados en:”,
font=(“Arial”, 9), fg=”#555555”, bg=“white”).pack()
ruta_display = REPORTES_PATH if len(REPORTES_PATH) <= 55 else “…” + REPORTES_PATH[-52:]
tk.Label(win, text=ruta_display,
font=(“Arial”, 9, “bold”), fg=”#1F4E78”, bg=“white”,
wraplength=360, justify=“center”).pack(pady=(2, 16))
frame_btns = tk.Frame(win, bg=“white”)
frame_btns.pack()
tk.Button(frame_btns, text=“Abrir carpeta”,
bg=”#1F4E78”, fg=“white”, font=(“Arial”, 10, “bold”),
width=15, relief=“flat”, cursor=“hand2”,
command=lambda: [os.startfile(REPORTES_PATH), win.destroy()]
).pack(side=“left”, padx=(0, 10))
tk.Button(frame_btns, text=“Aceptar”,
bg=”#E0E0E0”, fg=”#333333”, font=(“Arial”, 10),
width=10, relief=“flat”, cursor=“hand2”,
command=win.destroy).pack(side=“left”)

# ———————––

# CAMBIO 1: dias_credito_excel_like — fallback a Condiciones de pago MD

# ———————––

def dias_credito_excel_like(cond_pago, fecha_pago_dt, fe_contab_dt, cond_pago_md=””):
s = “” if pd.isna(cond_pago) else str(cond_pago).strip()
if s:
solo_digitos = re.sub(r”[^0-9]”, “”, s)
val = int(solo_digitos) if solo_digitos else 0
if val > 0:
return val
s_md = “” if pd.isna(cond_pago_md) else str(cond_pago_md).strip()
if s_md:
solo_digitos_md = re.sub(r”[^0-9]”, “”, s_md)
val_md = int(solo_digitos_md) if solo_digitos_md else 0
if val_md > 0:
return val_md
return None

def ejecutar_sap_flujo(fecha_sap: Optional[str] = None):
try:
import pythoncom
pythoncom.CoInitialize()
for f in [FILE_TC, FILE_RP]:
try:
if os.path.exists(f):
os.remove(f)
except Exception:
pass

```
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
    except Exception:
        try:
            SapGuiAuto = win32com.client.Dispatch("SapROTWr.SapROTWrapper")
            SapGuiAuto = SapGuiAuto.GetROTEntry("SAPGUI")
        except Exception as ex1:
            raise RuntimeError(
                f"No se pudo conectar a SAP GUI.\n"
                f"Asegúrate de que SAP esté abierto y con sesión iniciada.\n\n"
                f"Detalle: {ex1}"
            )
    application = SapGuiAuto.GetScriptingEngine
    if application is None:
        raise RuntimeError("SAP GUI Scripting no está habilitado.")
    try:
        connection = application.Children(0)
    except Exception as ex2:
        raise RuntimeError(f"No hay conexiones activas en SAP.\nDetalle: {ex2}")
    try:
        session = connection.Children(0)
    except Exception as ex3:
        raise RuntimeError(f"No hay sesiones activas en SAP.\nDetalle: {ex3}")

    ayer_sap = (pd.Timestamp.today() - pd.Timedelta(days=1)).strftime("%d.%m.%Y")
    if fecha_sap:
        fecha = pd.to_datetime(fecha_sap, format="%d/%m/%Y").strftime("%d.%m.%Y")
    else:
        fecha = ayer_sap

    print(f"SAP: Extrayendo Tipo de Cambio (TC.txt) para fecha {fecha}...")
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZF1FIR012"
    session.findById("wnd[0]").sendVKey(0)
    fecha_dt  = pd.to_datetime(fecha, format="%d.%m.%Y")
    fecha_low = (fecha_dt - pd.Timedelta(days=30)).strftime("%d.%m.%Y")
    session.findById("wnd[0]/usr/txtZDATE1-LOW").text = fecha_low
    session.findById("wnd[0]/usr/txtZDATE1-HIGH").text = fecha
    session.findById("wnd[0]/usr/ctxtSP$00001-LOW").text = "m"
    session.findById("wnd[0]/usr/ctxtSP$00002-LOW").text = "USD"
    session.findById("wnd[0]/usr/ctxtSP$00003-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtSP$00003-LOW").caretPosition = 0
    session.findById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "MXN"
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "JPY"
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "EUR"
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "BRL"
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "ARS"
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "COP"
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    grid = session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell")
    grid.pressToolbarContextButton("&MB_EXPORT")
    grid.selectContextMenuItem("&PC")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = REPORTES_PATH
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "TC.txt"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    if not esperar_archivo(FILE_TC, timeout=60, min_size=50):
        raise RuntimeError("TC.txt no se generó o quedó vacío.")

    print("SAP: Extrayendo Reporte de Pagos (RP.txt)...")
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nFBL1N"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(2)

    session.findById("wnd[0]/usr/chkX_MERK").selected = True
    session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").text = "MX10"
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = "/PAGOS_cv"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = REPORTES_PATH
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "RP.txt"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    if not esperar_archivo(FILE_RP, timeout=90, min_size=200):
        raise RuntimeError("RP.txt no se generó o quedó vacío.")

    return True

except Exception as e:
    mostrar_error("Error SAP", e)
    return False
```

def construir_tasas_usd(df_tc: pd.DataFrame) -> dict:
tasas = {“USD”: 1.0}
if df_tc.empty:
return tasas
cols_norm = {c: normalizar_col(c) for c in df_tc.columns}
col_rate = next((c for c, n in cols_norm.items() if “tipo cambio” in n or “tipo c” in n), None)
col_de = next((c for c, n in cols_norm.items() if n == “de”), None)
col_a = next((c for c, n in cols_norm.items() if n == “a”), None)
col_fecha = next((c for c, n in cols_norm.items() if “valido” in n or “valid” in n or “fecha” in n), None)
if not col_rate:
col_rate = df_tc.columns[0]
if not col_de or not col_a:
if len(df_tc.columns) >= 3:
col_de = col_de or df_tc.columns[1]
col_a = col_a or df_tc.columns[2]
if not col_de or not col_a:
return tasas
if col_fecha:
df_tc = df_tc.copy()
df_tc[”_fecha_dt”] = pd.to_datetime(df_tc[col_fecha], format=”%d.%m.%Y”, errors=“coerce”)
df_tc = df_tc.sort_values(”_fecha_dt”, ascending=True)
for _, r in df_tc.iterrows():
de = str(r.get(col_de, “”)).upper().strip()
a = str(r.get(col_a, “”)).upper().strip()
rate = limpiar_monto_sap(r.get(col_rate, 0))
if de == “USD” and a and rate > 0:
tasas[a] = rate
return tasas

def mapear_columnas_rp(df_rp: pd.DataFrame) -> pd.DataFrame:
cols_norm = {c: normalizar_col(c) for c in df_rp.columns}

```
def find_col(pred):
    for orig, n in cols_norm.items():
        if pred(n):
            return orig
    return None

col_doc = find_col(lambda n: n in ("n\u00ba doc.", "n\u00b0 doc.", "no doc.", "n doc.") or
                   ("doc" in n and n.startswith("n")))
if col_doc is None:
    col_doc = find_col(lambda n: "doc" in n and ("n" in n or "numero" in n))

col_fecontab    = find_col(lambda n: "contab" in n or n == "fe.contab.")
col_cond        = find_col(lambda n: n in ("cpag",) or ("cond" in n and "pago" in n))
col_fdoc        = find_col(lambda n: n in ("fecha doc.",) or
                           (n.startswith("fecha") and "doc" in n and "pago" not in n))
col_fpago       = find_col(lambda n: n in ("fecha pago",) or
                           (n.startswith("fecha") and "pago" in n))
col_imp_md      = find_col(lambda n: n in ("importe en md",) or
                           ("importe" in n and " md" in n))
col_mon         = find_col(lambda n: n in ("mon.", "mon"))
col_imp_ml      = find_col(lambda n: n in ("importe en ml",) or
                           ("importe" in n and " ml" in n))
col_mon_ml      = find_col(lambda n: n in ("ml", "moneda local"))
col_bp          = find_col(lambda n: n in ("bp", "bloqueo de pago", "bloqueo"))
col_ref         = find_col(lambda n: "refer" in n)
col_nombre      = find_col(lambda n: "nombre" in n or "acreedor" in n or "cliente" in n)
col_banco       = find_col(lambda n: "bco" in n or "banco" in n)
col_clase       = find_col(lambda n: "clase" in n)
col_doc_compras = find_col(lambda n: "doc.compr" in n or "documento compras" in n)
col_doc_comp    = find_col(lambda n: ("doc.comp" in n or "compens" in n) and "compr" not in n)
col_demora      = find_col(lambda n: "demora" in n)

rename = {}
if col_doc:         rename[col_doc]         = DOC_COL_CANON
if col_fecontab:    rename[col_fecontab]    = "Fe.contabilización"
if col_cond:        rename[col_cond]        = "Condiciones de pago"
if col_fdoc:        rename[col_fdoc]        = "Fecha de documento"
if col_fpago:       rename[col_fpago]       = "Fecha de pago"
if col_imp_md:      rename[col_imp_md]      = "Importe en moneda doc."
if col_mon:         rename[col_mon]         = "Moneda del documento"
if col_bp:          rename[col_bp]          = "Bloqueo de pago"
if col_ref:         rename[col_ref]         = "Referencia"
if col_nombre:      rename[col_nombre]      = "Nombre Cliente/Acreedor/Cuenta"
if col_banco:       rename[col_banco]       = "Banco propio"
if col_imp_ml:      rename[col_imp_ml]      = "Importe en moneda local"
if col_mon_ml:      rename[col_mon_ml]      = "Moneda local"
if col_clase:       rename[col_clase]       = "Clase de documento"
if col_doc_compras: rename[col_doc_compras] = "Documento compras"
if col_doc_comp:    rename[col_doc_comp]    = "Doc.compensación"
if col_demora:      rename[col_demora]      = "Demora tras vencimiento neto"

return df_rp.rename(columns=rename)
```

def procesar_informacion():
try:
print(“Python: Procesando archivos…”)

```
    if not os.path.exists(FILE_TC) or not os.path.exists(FILE_RP):
        raise RuntimeError("No encuentro TC.txt o RP.txt en la carpeta configurada.")

    df_tc = leer_tabla_sap_pipe(FILE_TC)
    df_rp_raw = leer_tabla_sap_pipe(FILE_RP)

    if df_tc.empty:
        raise RuntimeError("No se pudo leer TC.txt.")
    if df_rp_raw.empty:
        raise RuntimeError("No se pudo leer RP.txt.")

    df = mapear_columnas_rp(df_rp_raw)

    required = ["Cuenta", "Fe.contabilización", "Condiciones de pago",
                "Importe en moneda doc.", "Moneda del documento", DOC_COL_CANON]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise RuntimeError(f"RP.txt NO trae columnas requeridas: {missing}\nColumnas detectadas: {list(df.columns)}")

    tasas = construir_tasas_usd(df_tc)

    if "Importe en moneda local" in df.columns:
        df["Importe en moneda local"] = df["Importe en moneda local"].apply(limpiar_monto_sap)
    else:
        df["Importe en moneda local"] = 0.0

    df["Importe en moneda doc."] = df["Importe en moneda doc."].apply(limpiar_monto_sap)
    df["Moneda del documento"] = df["Moneda del documento"].astype(str).str.upper().str.strip()

    def a_usd(row):
        mon = row["Moneda del documento"]
        imp = row["Importe en moneda doc."]
        imp_ml = row.get("Importe en moneda local", 0.0)
        if mon == "USD":
            return round(imp, 2)
        if mon == "MXN":
            rate_mxn = tasas.get("MXN")
            if rate_mxn and rate_mxn != 0:
                return round(imp_ml / rate_mxn, 2)
            return 0.0
        if mon == "ARS":
            rate = tasas.get("ARS")
            if rate and rate != 0:
                return round(imp / rate, 2)
            return 0.0
        if mon == "EUR":
            rate = tasas.get("EUR")
            if rate and rate != 0:
                return round(imp / rate, 2)
            return 0.0
        if mon == "BRL":
            rate = tasas.get("BRL")
            if rate and rate != 0:
                return round(imp / rate, 2)
            return 0.0
        if mon == "JPY":
            rate = tasas.get("JPY")
            if rate and rate != 0:
                return round(imp / (rate * 100), 2)
            return 0.0
        if mon == "COP":
            rate = tasas.get("COP")
            if rate and rate != 0:
                return round(imp / (rate * 1000), 2)
            return 0.0
        rate = tasas.get(mon)
        if rate and rate != 0:
            return round(imp / rate, 2)
        return 0.0

    df["Total USD"] = df.apply(a_usd, axis=1).abs().astype(float)

    hoy = pd.Timestamp.today().normalize()
    sabado = hoy + pd.Timedelta(days=(5 - hoy.weekday()) % 7)
    semana_actual = int(hoy.isocalendar().week)

    def parse_fecha(serie):
        resultado = pd.to_datetime(serie, format="%d.%m.%Y", errors="coerce")
        mask = resultado.isna()
        if mask.any():
            resultado[mask] = pd.to_datetime(serie[mask], format="%d/%m/%Y", errors="coerce")
        return resultado

    df["Fe.contabilización_dt"] = parse_fecha(df["Fe.contabilización"])
    df["Fecha de documento_dt"]  = parse_fecha(
        df.get("Fecha de documento", pd.Series([pd.NaT] * len(df))))
    df["Fecha de pago_dt"] = parse_fecha(
        df.get("Fecha de pago", pd.Series([pd.NaT] * len(df))))

    def semana_excel(serie):
        """Calcula número de semana igual que NUM.DE.SEMANA de Excel (modo 1, semana empieza el domingo)"""
        return serie.apply(lambda d: int(d.strftime("%U")) + 1 if pd.notna(d) else 0)

    df["Adeudos"] = 0.0
    df["Terceros / Especiales"] = "Terceros"
    df["Descripción"]           = ""
    df["Criticidad"]            = ""
    df["Naturaleza 1"]          = ""
    df["Naturaleza 2"]          = ""
    # CAMBIO 2: inicializar Condiciones de pago MD antes del catálogo
    df["Condiciones de pago MD"] = ""
    for _et in range(1, 11):
        df[f"Etiqueta {_et}"] = ""

    global CATALOGO_PATH
    if not os.path.exists(CATALOGO_PATH):
        from tkinter import filedialog
        messagebox.showinfo(
            "Catálogo no encontrado",
            "No se encontró el catálogo.\nSe abrirá el explorador para que lo selecciones."
        )
        ruta_manual = filedialog.askopenfilename(
            title="Selecciona el catálogo Excel",
            initialdir=CATALOGOS_PATH,
            filetypes=[("Excel", "*.xlsx *.xls")]
        )
        if ruta_manual:
            CATALOGO_PATH = ruta_manual
        else:
            messagebox.showwarning("Proceso cancelado", "No se seleccionó un catálogo. El proceso fue cancelado.")
            return

    if os.path.exists(CATALOGO_PATH):
        df_cat = pd.read_excel(CATALOGO_PATH)
        df_cat.columns = [normalizar_col(c) for c in df_cat.columns]

        if "clave" not in df_cat.columns:
            raise RuntimeError(f"El catálogo no trae columna 'Clave'. Columnas: {df_cat.columns.tolist()}")

        df["key"]     = df["Cuenta"].apply(norm_key)
        df_cat["key"] = df_cat["clave"].apply(norm_key)
        df_cat        = df_cat.drop_duplicates(subset=["key"], keep="first")

        col_nat1 = "para pagos ( l1) naturaleza 1"
        col_nat2 = "categoria de gasto (l2) naturaleza 2"

        if col_nat1 not in df_cat.columns:
            raise RuntimeError(f"No encuentro Naturaleza 1 en catálogo. Buscaba '{col_nat1}'.")
        if col_nat2 not in df_cat.columns:
            raise RuntimeError(f"No encuentro Naturaleza 2 en catálogo. Buscaba '{col_nat2}'.")

        map_nat1 = dict(zip(df_cat["key"], df_cat[col_nat1]))
        map_nat2 = dict(zip(df_cat["key"], df_cat[col_nat2]))

        if "especial/ no" in df_cat.columns:
            df["Terceros / Especiales"] = df["key"].map(dict(zip(df_cat["key"], df_cat["especial/ no"]))).fillna("Terceros")
        if "criticidad" in df_cat.columns:
            df["Criticidad"] = df["key"].map(dict(zip(df_cat["key"], df_cat["criticidad"]))).fillna("")
        if "descripcion" in df_cat.columns:
            df["Descripción"] = df["key"].map(dict(zip(df_cat["key"], df_cat["descripcion"]))).fillna("")

        def naturaleza1_logic(fecha_doc_dt, terceros_val, key):
            try:
                if pd.isna(fecha_doc_dt):
                    return ""
                if int(fecha_doc_dt.year) < 2025:
                    return "- Manor 024"
                if str(terceros_val).strip().upper() == "EMPLEADOS":
                    return "Nomina y Personas"
                v = map_nat1.get(key, "")
                if v is None or (isinstance(v, float) and np.isnan(v)):
                    return ""
                return str(v)
            except Exception:
                return ""

        df["Naturaleza 1"] = df.apply(
            lambda r: naturaleza1_logic(r["Fecha de documento_dt"], r["Terceros / Especiales"], r["key"]), axis=1)
        df["Naturaleza 2"] = df["key"].map(map_nat2).fillna("")
        for _et in range(1, 11):
            col_et = f"etiqueta {_et}"
            if col_et in df_cat.columns:
                df[f"Etiqueta {_et}"] = df["key"].map(dict(zip(df_cat["key"], df_cat[col_et]))).fillna("")
        # CAMBIO 2: mapear Condiciones de pago MD desde catálogo
        col_cond_md = next((c for c in df_cat.columns if "dias de credito md" in c or "días de crédito md" in c or "condiciones de pago md" in c), None)
        if col_cond_md:
            df["Condiciones de pago MD"] = df["key"].map(dict(zip(df_cat["key"], df_cat[col_cond_md]))).fillna("")
        match_rate = df["key"].isin(df_cat["key"]).mean()
        print(f"[CAT] Match rate (Cuenta vs Clave): {match_rate:.2%}")
    else:
        print("[CAT] No existe el archivo de catálogo. Se dejan campos por defecto.")

    # CAMBIO 3: Días de crédito DESPUÉS del catálogo para poder usar Condiciones de pago MD
    df["Días de crédito"] = df.apply(
        lambda r: dias_credito_excel_like(
            r.get("Condiciones de pago", ""),
            r.get("Fecha de pago_dt", pd.NaT),
            r.get("Fe.contabilización_dt", pd.NaT),
            r.get("Condiciones de pago MD", ""),
        ), axis=1
    ).apply(lambda x: int(x) if x is not None else None).apply(lambda x: int(x) if pd.notna(x) else None)

    df["Fecha_FF_Dt"]         = df["Fecha de documento_dt"] + pd.to_timedelta(df["Días de crédito"].fillna(0), unit="D")
    df["Fecha_Min_Dt"]        = df["Fe.contabilización_dt"] + pd.to_timedelta(15, unit="D")
    df["Fecha_Venc_Final_Dt"] = df[["Fecha_FF_Dt", "Fecha_Min_Dt"]].max(axis=1)

    df["Estatus BI (Hoy)"] = np.where((df["Fecha_Venc_Final_Dt"] - hoy).dt.days < 0, "Vencido", "Por vencer")

    df["Estatus BI"] = np.where(
        (hoy - df["Fecha_Venc_Final_Dt"]).dt.days < 0, "Por vencer", "Vencido"
    )

    df["Estatus Proveedor (Hoy)"] = np.where(
        (hoy - df["Fecha_FF_Dt"]).dt.days < 0, "Por vencer", "Vencido"
    )

    df["Estatus Proveedor"] = np.where(
        (hoy - df["Fecha_FF_Dt"]).dt.days < 0, "Por vencer", "Vencido"
    )

    df["Semana Pago por Vencimiento"] = semana_excel(df["Fecha_Venc_Final_Dt"])
    df["Sem registro"]               = semana_excel(df["Fe.contabilización_dt"])
    df["Semana Vencimiento SAP"]     = semana_excel(df["Fecha de pago_dt"])

    df["Días de retraso (Hoy)"] = np.where(
        df["Estatus BI (Hoy)"] == "Vencido",
        (hoy - df["Fecha_FF_Dt"]).dt.days,
        0
    ).astype(int)

    df["Reporte días vencidos (Hoy)"] = np.select(
        [(df["Días de retraso (Hoy)"] <= 0),
         (df["Días de retraso (Hoy)"] <= 15),
         (df["Días de retraso (Hoy)"] <= 30),
         (df["Días de retraso (Hoy)"] <= 60)],
        [df["Semana Pago por Vencimiento"], "<15D", "+15D", "+30D"],
        default="+60D"
    )

    df["Días de retraso"] = np.where(
        df["Estatus BI"] == "Vencido",
        (sabado - df["Fecha_FF_Dt"]).dt.days,
        0
    ).astype(int)

    df["Reporte días vencidos"] = np.select(
        [(df["Días de retraso"] <= 0),
         (df["Días de retraso"] <= 15),
         (df["Días de retraso"] <= 30),
         (df["Días de retraso"] <= 60)],
        [df["Semana Pago por Vencimiento"], "<15D", "+15D", "+30D"],
        default="+60D"
    )

    df["Semana propuesta pago"] = ""
    df["Día de pago"]           = ""

    df["Semana Vencimiento Real F doc"] = semana_excel(df["Fecha_FF_Dt"])
    df["Semana Venc por Proveedor"]     = semana_excel(df["Fecha_FF_Dt"])

    df["VS SAP"] = np.where(
        df["Fecha de pago_dt"].isna(), "",
        (df["Fecha de pago_dt"] - df["Fecha_Venc_Final_Dt"]).dt.days
    )

    for col_doc in [DOC_COL_CANON, "Documento compras", "Doc.compensación"]:
        if col_doc in df.columns:
            s = df[col_doc].astype(str).str.replace(r"\.0$", "", regex=True)
            s = s.replace("nan", "").replace("NaT", "")
            df[col_doc] = s

    df["Fecha de Pago por F.F."]           = df["Fecha_FF_Dt"].dt.strftime("%d/%m/%Y")
    df["Fecha Minima por contabilización"]  = df["Fecha_Min_Dt"].dt.strftime("%d/%m/%Y")
    df["Fecha Vencimiento"]                 = df["Fecha_Venc_Final_Dt"].dt.strftime("%d/%m/%Y")
    df["Vencimiento Diario (Hoy)"]          = hoy.strftime("%d/%m/%Y")
    df["Fecha Revisión"]                    = sabado.strftime("%d/%m/%Y")
    df["Fe.contabilización"]                = df["Fe.contabilización_dt"].dt.strftime("%d/%m/%Y")
    df["Fecha de documento"]                = df["Fecha de documento_dt"].dt.strftime("%d/%m/%Y")
    df["Fecha de pago"]                     = df["Fecha de pago_dt"].dt.strftime("%d/%m/%Y")

    df["Corrección en SAP"] = np.where(
        df["Fecha_Venc_Final_Dt"] >= df["Fecha de pago_dt"],
        "Corregir Fecha PAGO", ""
    )

    # CAMBIO 2: Condiciones de pago MD después de Condiciones de pago en columnas_ordenadas
    columnas_ordenadas = [
        "Cuenta", "Bloqueo de pago", "Banco propio", "Fe.contabilización", DOC_COL_CANON,
        "Condiciones de pago", "Condiciones de pago MD", "Nombre Cliente/Acreedor/Cuenta", "Referencia",
        "Fecha de documento",
        "Importe en moneda local", "Moneda local", "Importe en moneda doc.", "Moneda del documento",
        "Fecha de pago", "Clase de documento", "Documento compras", "Demora tras vencimiento neto",
        "Doc.compensación", "Días de crédito", "Sem registro", "Semana Vencimiento SAP",
        "Total USD", "Semana Pago por Vencimiento", "Vencimiento Diario (Hoy)",
        "Corrección en SAP", "VS SAP", "Fecha Revisión", "Terceros / Especiales", "Descripción",
        "Naturaleza 1", "Naturaleza 2", "Naturaleza 3", "Etiqueta 1", "Etiqueta 2", "Etiqueta 3", "Etiqueta 4", "Etiqueta 5",
        "Etiqueta 6", "Etiqueta 7", "Etiqueta 8", "Etiqueta 9", "Etiqueta 10", "Criticidad",
        "Fecha de Pago por F.F.", "Fecha Minima por contabilización", "Fecha Vencimiento",
        "Semana Vencimiento Real F doc", "Semana Venc por Proveedor",
        "Estatus BI (Hoy)", "Estatus BI", "Estatus Proveedor (Hoy)", "Estatus Proveedor",
        "Días de retraso (Hoy)", "Reporte días vencidos (Hoy)",
        "Días de retraso", "Reporte días vencidos", "Adeudos", "Semana propuesta pago", "Día de pago"
    ]

    for c in columnas_ordenadas:
        if c not in df.columns:
            df[c] = ""

    df_final = df.reindex(columns=columnas_ordenadas).copy()

    PROPUESTA_PATH = os.path.join(CATALOGOS_PATH, "propuesta_pago.xlsx")
    if os.path.exists(PROPUESTA_PATH):
        df_prop = pd.read_excel(PROPUESTA_PATH)
        col_doc_prop = next((c for c in df_prop.columns if "documento" in str(c).lower() or "doc" in str(c).lower()), None)
        col_sem_prop = next((c for c in df_prop.columns if "semana" in str(c).lower()), None)
        col_dia_prop = next((c for c in df_prop.columns if "día de pago" in str(c).lower() or "dia de pago" in str(c).lower()), None)
        col_nat3_prop = next((c for c in df_prop.columns if "naturaleza 3" in str(c).lower()), None)
        if col_doc_prop:
            df_prop["_doc_key"] = df_prop[col_doc_prop].astype(str).str.strip().str.zfill(ZFILL_LEN)
            df_final["_doc_key"] = df_final[DOC_COL_CANON].astype(str).str.strip().str.zfill(ZFILL_LEN)
            print(f"[CRUCE] Muestra doc propuesta: {df_prop['_doc_key'].head(3).tolist()}")
            print(f"[CRUCE] Muestra doc base: {df_final['_doc_key'].head(3).tolist()}")
            if col_sem_prop:
                mapa_sem = dict(zip(df_prop["_doc_key"], df_prop[col_sem_prop].astype(str)))
                df_final["Semana propuesta pago"] = df_final["_doc_key"].map(mapa_sem).fillna("")
                df_final["Semana propuesta pago"] = df_final["Semana propuesta pago"].apply(
                    lambda x: str(int(float(x))) if x not in ("", "nan") and x.replace(".", "").isdigit() else x
                )
                print(f"[CRUCE] Matches semana: {(df_final['Semana propuesta pago'] != '').sum()}")
            if col_dia_prop:
                mapa_dia = dict(zip(df_prop["_doc_key"], df_prop[col_dia_prop].astype(str)))
                df_final["Día de pago"] = df_final["_doc_key"].map(mapa_dia).fillna("")
                df_final["Día de pago"] = df_final["Día de pago"].replace("nan", "")
            if col_nat3_prop:
                mapa_nat3 = dict(zip(df_prop["_doc_key"], df_prop[col_nat3_prop].astype(str)))
                df_final["Naturaleza 3"] = df_final["_doc_key"].map(mapa_nat3).fillna("")
                df_final["Naturaleza 3"] = df_final["Naturaleza 3"].replace("nan", "")
            df_final.drop(columns=["_doc_key"], inplace=True)

    semana_sabado = semana_excel(pd.Series([sabado]))[0]
    sem_prop = pd.to_numeric(df_final["Semana propuesta pago"].replace("", np.nan), errors="coerce")
    total_usd = pd.to_numeric(df_final["Total USD"], errors="coerce").fillna(0.0)
    sem_venc_prov = df_final["Semana propuesta pago"].apply(lambda x: float(x) if str(x).strip() not in ("", "nan") else np.nan)
    semana_hoy = semana_excel(pd.Series([hoy]))[0]
    df_final["Adeudos"] = np.where(
        sem_venc_prov.isna(), total_usd,
        np.where(sem_venc_prov <= semana_hoy, np.nan, total_usd)
    )

    for col_num in ["Importe en moneda local", "Importe en moneda doc.", "Total USD"]:
        if col_num in df_final.columns:
            df_final[col_num] = pd.to_numeric(df_final[col_num], errors="coerce").fillna(0.0)

    for col_txt in df_final.columns:
        if col_txt not in ["Importe en moneda local", "Importe en moneda doc.", "Total USD", "Adeudos"]:
            df_final[col_txt] = df_final[col_txt].astype(str).replace("nan", "").replace("NaT", "")

    mapeo_resumen = {
        DOC_COL_CANON:                    DOC_COL_CANON,
        "Naturaleza 1":                   "Naturaleza 1",
        "Cuenta":                         "# Proveedor",
        "Referencia":                     "Referencia",
        "Nombre Cliente/Acreedor/Cuenta": "Nombre",
        "Total USD":                      "Importe USD",
        "Reporte días vencidos (Hoy)":    "Rango Vencimiento",
        "Semana Pago por Vencimiento":    "Semana Vencimiento",
        "Estatus BI":                     "Estatus",
    }
    cols_a_extraer = [c for c in mapeo_resumen.keys() if c in df_final.columns]
    df_resumen = df_final[cols_a_extraer].rename(columns=mapeo_resumen).copy()

    idx_nat1 = df_resumen.columns.get_loc("Naturaleza 1") + 1
    df_resumen.insert(idx_nat1, "Naturaleza 3", df_final["Naturaleza 3"].values)
    for _et in range(1, 11):
        df_resumen.insert(df_resumen.columns.get_loc("Naturaleza 3") + _et, f"Etiqueta {_et}", df_final[f"Etiqueta {_et}"].values if f"Etiqueta {_et}" in df_final.columns else "")

    df_resumen.loc[df_resumen["Estatus"] == "Vencido", "Semana Vencimiento"] = ""
    df_resumen.loc[df_resumen["Estatus"] == "Por vencer", "Rango Vencimiento"] = ""
    df_resumen["Semana propuesta pago"] = df_final["Semana propuesta pago"].values
    df_resumen["Día de pago"] = df_final["Día de pago"].values

    with pd.ExcelWriter(RESUMEN_PATH, engine="xlsxwriter") as writer_res:
        df_resumen.to_excel(writer_res, index=False, sheet_name="Resumen", startrow=5)
        wb_res       = writer_res.book
        ws_res       = writer_res.sheets["Resumen"]
        fmt_azul_res = wb_res.add_format({"bg_color": "#1F4E78"})
        fmt_tit_res  = wb_res.add_format({"bold": True, "font_size": 16, "font_color": "white", "bg_color": "#1F4E78"})
        fmt_res_hdr  = wb_res.add_format({"bold": True, "bg_color": "#BFBFBF", "border": 1, "align": "center"})
        fmt_res_cur  = wb_res.add_format({"num_format": "#,##0.00", "border": 1})
        fmt_res_std  = wb_res.add_format({"border": 1})
        fmt_vencido  = wb_res.add_format({"bg_color": "#FF0000", "font_color": "white", "bold": True, "border": 1})
        fmt_porvencer = wb_res.add_format({"bg_color": "#00B050", "font_color": "white", "bold": True, "border": 1})
        fmt_blanco   = wb_res.add_format({"border": 0, "bg_color": "#FFFFFF"})
        for r in range(5):
            ws_res.set_row(r, 15, fmt_blanco)
        for col_num, value in enumerate(df_resumen.columns.values):
            ws_res.write(5, col_num, value, fmt_res_hdr)
            if value == "Importe USD":
                ws_res.set_column(col_num, col_num, 15, fmt_res_cur)
            else:
                ws_res.set_column(col_num, col_num, 25, fmt_res_std)

        col_estatus = list(df_resumen.columns).index("Estatus") if "Estatus" in df_resumen.columns else None
        if col_estatus is not None:
            for row_num, valor in enumerate(df_resumen["Estatus"]):
                fmt = fmt_vencido if str(valor).strip() == "Vencido" else fmt_porvencer
                ws_res.write(row_num + 6, col_estatus, valor, fmt)

    with pd.ExcelWriter(OUTPUT_PATH, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Reporte", startrow=5)
        workbook  = writer.book
        worksheet = writer.sheets["Reporte"]
        fmt_azul  = workbook.add_format({"bg_color": "#1F4E78"})
        fmt_hdr   = workbook.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white", "border": 1, "align": "center"})
        fmt_money = workbook.add_format({"num_format": "#,##0.00", "border": 1})
        fmt_txt   = workbook.add_format({"border": 1})
        for r in range(5):
            worksheet.set_row(r, 25, fmt_azul)
        worksheet.write("D2", "REPORTE DE PAGOS BRASKEM",
                        workbook.add_format({"bold": True, "font_size": 18, "font_color": "white", "bg_color": "#1F4E78"}))
        worksheet.write("D3", f"Generado: {hoy.strftime('%d/%m/%Y')}",
                        workbook.add_format({"font_color": "white", "bg_color": "#1F4E78"}))
        cols_grupo1 = {
            "Cuenta", "Bloqueo de pago", "Banco propio", "Fe.contabilización",
            "N° documento", "Condiciones de pago", "Condiciones de pago MD", "Nombre Cliente/Acreedor/Cuenta",
            "Referencia", "Fecha de documento", "Importe en moneda local", "Moneda local",
            "Importe en moneda doc.", "Moneda del documento", "Fecha de pago",
            "Clase de documento", "Documento compras", "Demora tras vencimiento neto",
            "Doc.compensación"
        }
        cols_grupo2 = {
            "Días de crédito", "Sem registro", "Vencimiento Diario (Hoy)", "Fecha Revisión",
            "Fecha de Pago por F.F.", "Fecha Minima por contabilización", "Fecha Vencimiento",
            "Semana Venc por Proveedor", "Semana Pago por Vencimiento", "Semana Vencimiento SAP",
            "Semana Vencimiento Real F doc", "Total USD", "Estatus BI (Hoy)", "Estatus BI",
            "Estatus Proveedor (Hoy)", "Estatus Proveedor", "Corrección en SAP", "VS SAP",
            "Días de retraso (Hoy)", "Reporte días vencidos (Hoy)",
            "Días de retraso", "Reporte días vencidos", "Adeudos"
        }
        cols_grupo3 = {
            "Terceros / Especiales", "Naturaleza 1", "Criticidad", "Descripción", "Naturaleza 2",
            "Condiciones de pago MD"
        }
        cols_grupo4 = {
            "Semana propuesta pago", "Día de pago", "Naturaleza 3"
        }
        fmt_hdr_g1 = workbook.add_format({"bold": True, "bg_color": "#BFBFBF", "font_color": "#1A1A1A", "border": 1, "align": "center"})
        fmt_hdr_g2 = workbook.add_format({"bold": True, "bg_color": "#7B5EA7", "font_color": "white",   "border": 1, "align": "center"})
        fmt_hdr_g3 = workbook.add_format({"bold": True, "bg_color": "#4E8B5F", "font_color": "white",   "border": 1, "align": "center"})
        fmt_hdr_g4 = workbook.add_format({"bold": True, "bg_color": "#7BA7C7", "font_color": "white",   "border": 1, "align": "center"})
        for i, col in enumerate(columnas_ordenadas):
            if col in cols_grupo1:
                hdr_fmt = fmt_hdr_g1
            elif col in cols_grupo2:
                hdr_fmt = fmt_hdr_g2
            elif col in cols_grupo3:
                hdr_fmt = fmt_hdr_g3
            elif col in cols_grupo4:
                hdr_fmt = fmt_hdr_g4
            else:
                hdr_fmt = fmt_hdr
            worksheet.write(5, i, col, hdr_fmt)
            if col in ["Total USD", "Adeudos", "Importe en moneda local", "Importe en moneda doc."]:
                worksheet.set_column(i, i, 18, fmt_money)
            else:
                worksheet.set_column(i, i, 18, fmt_txt)

    REPORTE2_PATH = os.path.join(REPORTES_PATH, "Reporte Tesoreria.xlsx")
    cols_reporte2 = [
        "Cuenta", "Bloqueo de pago", "Banco propio", "Fe.contabilización", DOC_COL_CANON,
        "Condiciones de pago", "Condiciones de pago MD", "Nombre Cliente/Acreedor/Cuenta", "Referencia",
        "Fecha de documento", "Importe en moneda local", "Moneda local",
        "Importe en moneda doc.", "Moneda del documento", "Fecha de pago",
        "Clase de documento", "Documento compras", "Demora tras vencimiento neto",
        "Doc.compensación", "Terceros / Especiales", "Naturaleza 1", "Naturaleza 2", "Naturaleza 3",
        "Etiqueta 1", "Etiqueta 2", "Etiqueta 3", "Etiqueta 4", "Etiqueta 5",
        "Etiqueta 6", "Etiqueta 7", "Etiqueta 8", "Etiqueta 9", "Etiqueta 10",
        "Semana Vencimiento SAP", "Semana Vencimiento Real F doc", "Semana Venc por Proveedor", "Total USD",
        "Estatus BI (Hoy)", "Estatus Proveedor (Hoy)", "Estatus BI", "Estatus Proveedor",
        "Fecha de Pago por F.F.", "Fecha Minima por contabilización", "Fecha Vencimiento",
        "Semana propuesta pago", "Día de pago"
    ]
    df_reporte2 = df_final.reindex(columns=[c for c in cols_reporte2 if c in df_final.columns]).copy()

    años_validos = [hoy.year, hoy.year - 1]
    if "Fe.contabilización_dt" in df.columns:
        mask = df["Fe.contabilización_dt"].dt.year.isin(años_validos)
        df_reporte2 = df_reporte2[mask.values]

    if "Importe en moneda local" in df_reporte2.columns:
        df_reporte2 = df_reporte2[pd.to_numeric(df_reporte2["Importe en moneda local"], errors="coerce") < 0]

    if "Doc.compensación" in df_reporte2.columns:
        df_reporte2 = df_reporte2[df_reporte2["Doc.compensación"].astype(str).str.strip().isin(["", "nan", "0"])]

    if "Clase de documento" in df_reporte2.columns:
        df_reporte2 = df_reporte2[~df_reporte2["Clase de documento"].astype(str).str.strip().isin(["TR", "RR", "KZ", "ZP"])]

    if "Bloqueo de pago" in df_reporte2.columns:
        df_reporte2 = df_reporte2[df_reporte2["Bloqueo de pago"].astype(str).str.strip() == "V"]

    for col_nuevo in ["Estatus BI (Hoy)", "Estatus Proveedor (Hoy)", "Estatus BI", "Estatus Proveedor"]:
        if col_nuevo not in df_reporte2.columns and col_nuevo in df.columns:
            df_reporte2[col_nuevo] = df[col_nuevo].values

    with pd.ExcelWriter(REPORTE2_PATH, engine="xlsxwriter") as writer2:
        df_reporte2.to_excel(writer2, index=False, sheet_name="Reporte", startrow=5)
        wb2  = writer2.book
        ws2  = writer2.sheets["Reporte"]
        fmt_hdr2   = wb2.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white", "border": 1, "align": "center"})
        fmt_azul2  = wb2.add_format({"bg_color": "#1F4E78"})
        fmt_money2 = wb2.add_format({"num_format": "#,##0.00", "border": 1})
        fmt_txt2   = wb2.add_format({"border": 1})
        for r in range(5):
            ws2.set_row(r, 25, fmt_azul2)
        ws2.write("D2", "REPORTE PROVEEDORES BRASKEM",
                  wb2.add_format({"bold": True, "font_size": 18, "font_color": "white", "bg_color": "#1F4E78"}))
        ws2.write("D3", f"Generado: {hoy.strftime('%d/%m/%Y')}",
                  wb2.add_format({"font_color": "white", "bg_color": "#1F4E78"}))
        for i, col in enumerate(df_reporte2.columns):
            ws2.write(5, i, col, fmt_hdr2)
            if col in ["Total USD", "Importe en moneda local", "Importe en moneda doc."]:
                ws2.set_column(i, i, 18, fmt_money2)
            else:
                ws2.set_column(i, i, 20, fmt_txt2)

    root.after(0, popup_exito)

except Exception as e:
    mostrar_error("Error de Datos", e)
```

from tkinter import ttk
import threading

C_AZUL_OSC  = “#0D2D6E”
C_AZUL_MED  = “#1F4E78”
C_AZUL_CLAR = “#2E75B6”
C_ACENTO    = “#00AEEF”
C_BG        = “#F4F7FB”
C_CARD      = “#FFFFFF”
C_TEXTO     = “#1A2B4A”
C_GRIS      = “#8A9AB5”
C_BORDE     = “#D6E4F0”
C_VERDE     = “#1A7A1A”
C_NARANJA   = “#E07B00”
C_ROJO      = “#CC0000”

root = tk.Tk()
root.title(“Braskem Automator v6.7”)
root.geometry(“460x520”)
root.configure(bg=C_BG)
root.resizable(False, False)

barra_top = tk.Frame(root, bg=C_AZUL_OSC, height=6)
barra_top.pack(fill=“x”, side=“top”)

card = tk.Frame(root, bg=C_CARD, bd=0, highlightthickness=1,
highlightbackground=C_BORDE)
card.pack(fill=“both”, expand=True, padx=24, pady=(18, 18))

logo_cargado = False
if PIL_DISPONIBLE:
try:
posibles_rutas = [
os.path.join(CATALOGOS_PATH, “logo.png”),
resource_path(“logo.png”),
r”C:\Users\migueh04\Documents\Reporte de pagos\logo.png”,
]
logo_file = next((r for r in posibles_rutas if os.path.exists(r)), None)
if logo_file:
img = Image.open(logo_file)
ancho_objetivo = 150
ratio = ancho_objetivo / img.width
alto_objetivo = int(img.height * ratio)
img = img.resize((ancho_objetivo, alto_objetivo), Image.LANCZOS)
logo_img = ImageTk.PhotoImage(img)
lbl_logo = tk.Label(card, image=logo_img, bg=C_CARD)
lbl_logo.image = logo_img
lbl_logo.pack(pady=(20, 4))
logo_cargado = True
except Exception as e:
print(f”Error cargando logo: {e}”)

if not logo_cargado:
tk.Label(card, text=“BRASKEM IDESA”, font=(“Helvetica”, 17, “bold”),
fg=C_AZUL_OSC, bg=C_CARD).pack(pady=(22, 4))

tk.Label(card, text=“Automatización de Reportes”,
font=(“Helvetica”, 12, “bold”), fg=C_AZUL_MED, bg=C_CARD).pack(pady=(2, 0))
tk.Label(card, text=“Módulo de Facturación y Pagos”,
font=(“Helvetica”, 9), fg=C_GRIS, bg=C_CARD).pack(pady=(0, 4))

sep_frame = tk.Frame(card, bg=C_CARD)
sep_frame.pack(fill=“x”, padx=28, pady=(6, 14))
tk.Frame(sep_frame, height=1, bg=C_BORDE).pack(fill=“x”)
tk.Frame(sep_frame, height=2, bg=C_ACENTO, width=60).place(x=0, y=0)

var_sap = tk.BooleanVar(value=EJECUTAR_SAP)
chk_sap = tk.Checkbutton(card, text=“Conectar a SAP”, variable=var_sap,
font=(“Helvetica”, 9), fg=C_TEXTO, bg=C_CARD,
activebackground=C_CARD, selectcolor=C_CARD, cursor=“hand2”)
chk_sap.pack(anchor=“w”, padx=28, pady=(0, 4))

frame_fecha_outer = tk.Frame(card, bg=C_CARD)
frame_fecha_outer.pack(pady=(0, 6))

tk.Label(frame_fecha_outer, text=“Fecha de corte”,
font=(“Helvetica”, 9, “bold”), fg=C_TEXTO, bg=C_CARD).pack(anchor=“w”, padx=28)

frame_fecha = tk.Frame(card, bg=C_BG, bd=0, highlightthickness=1,
highlightbackground=C_BORDE)
frame_fecha.pack(padx=28, pady=(2, 12), fill=“x”)

tk.Label(frame_fecha, text=“📅”, font=(“Helvetica”, 11),
bg=C_BG, fg=C_AZUL_MED).pack(side=“left”, padx=(10, 4), pady=8)
tk.Label(frame_fecha, text=“dd/mm/aaaa”,
font=(“Helvetica”, 8), fg=C_GRIS, bg=C_BG).pack(side=“right”, padx=(0, 10), pady=8)

ayer_default = pd.Timestamp.today().strftime(”%d/%m/%Y”)
entry_fecha = tk.Entry(frame_fecha, font=(“Helvetica”, 11), justify=“center”,
bd=0, relief=“flat”, bg=C_BG, fg=C_TEXTO,
insertbackground=C_AZUL_MED, width=12)
entry_fecha.insert(0, ayer_default)
entry_fecha.pack(side=“left”, pady=8, padx=(0, 4))

btn = tk.Button(
card,
text=“▶   Ejecutar Reporte”,
bg=C_AZUL_OSC,
fg=“white”,
font=(“Helvetica”, 11, “bold”),
relief=“flat”,
cursor=“hand2”,
activebackground=C_AZUL_CLAR,
activeforeground=“white”,
pady=12,
)
btn.pack(fill=“x”, padx=28, pady=(4, 8))

def on_enter(e): btn.config(bg=C_AZUL_CLAR)
def on_leave(e): btn.config(bg=C_AZUL_OSC if btn[“state”] == “normal” else “#888888”)
btn.bind(”<Enter>”, on_enter)
btn.bind(”<Leave>”, on_leave)

tk.Frame(card, height=1, bg=C_BORDE).pack(fill=“x”, padx=28, pady=(4, 8))

lbl_estado = tk.Label(card, text=“Listo para ejecutar”,
font=(“Helvetica”, 9), fg=C_GRIS, bg=C_CARD)
lbl_estado.pack(pady=(0, 4))

style = ttk.Style()
style.theme_use(“clam”)
style.configure(“Braskem.Horizontal.TProgressbar”,
troughcolor=C_BORDE, background=C_ACENTO,
thickness=4, borderwidth=0)
progress = ttk.Progressbar(card, mode=“indeterminate”, length=400,
style=“Braskem.Horizontal.TProgressbar”)

tk.Label(card, text=f”v6.7  •  {pd.Timestamp.today().strftime(’%Y’)}  •  Braskem Idesa”,
font=(“Helvetica”, 7), fg=C_BORDE, bg=C_CARD).pack(side=“bottom”, pady=(0, 10))

tk.Frame(root, bg=C_ACENTO, height=3).pack(fill=“x”, side=“bottom”)

def set_ui_running(running: bool):
if running:
btn.config(state=“disabled”, text=“⏳  Procesando…”, bg=”#888888”)
progress.pack(fill=“x”, padx=28, pady=(0, 6))
progress.start(10)
else:
progress.stop()
progress.pack_forget()
btn.config(state=“normal”, text=“▶   Ejecutar Reporte”, bg=C_AZUL_OSC)

def iniciar_todo():
fecha_usr = entry_fecha.get().strip()
try:
pd.to_datetime(fecha_usr, format=”%d/%m/%Y”)
except Exception:
messagebox.showerror(
“Fecha inválida”,
“La fecha “ + fecha_usr + “ no es valida.\nUsa el formato dd/mm/aaaa\nEjemplo: 21/02/2026”
)
return

```
def run():
    set_ui_running(True)
    try:
        if var_sap.get():
            lbl_estado.config(text="🔄  Conectando a SAP...", fg=C_NARANJA)
            ok = ejecutar_sap_flujo(fecha_sap=fecha_usr)
            if ok:
                lbl_estado.config(text="✅  SAP completado, procesando datos...", fg=C_NARANJA)
                procesar_informacion()
        else:
            lbl_estado.config(text="🔄  Procesando datos...", fg=C_NARANJA)
            procesar_informacion()
        lbl_estado.config(text="✅  Proceso completado exitosamente", fg=C_VERDE)
    except Exception as e:
        lbl_estado.config(text="❌  Error: " + str(e)[:45], fg=C_ROJO)
    finally:
        set_ui_running(False)

threading.Thread(target=run, daemon=True).start()
```

btn.config(command=iniciar_todo)

root.mainloop()