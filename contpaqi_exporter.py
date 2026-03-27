# ==========================
# contpaqi_export_pro.py (ULTRA PRO v5)
# Excel SolutionsV — Exportador de Nominas ContpaqI
# ==========================

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import threading
import subprocess
import re
import datetime
import pyodbc
import pandas as pd
import json
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

CONFIG_FILE = "config.json"

# ──────────────────────────────────────────────
# ODBC: drivers a intentar en orden de preferencia
# ──────────────────────────────────────────────
ODBC_DRIVERS_PRIORITY = [
    "ODBC Driver 17 for SQL Server",
    "ODBC Driver 18 for SQL Server",
    "ODBC Driver 13 for SQL Server",
    "SQL Server Native Client 11.0",
    "SQL Server Native Client 10.0",
    "SQL Server",                      # driver generico, siempre presente en Windows
]

def get_available_odbc_driver():
    """Retorna el primer driver ODBC disponible en el sistema, o None si no hay ninguno."""
    installed = [d for d in pyodbc.drivers()]
    for preferred in ODBC_DRIVERS_PRIORITY:
        if preferred in installed:
            return preferred
    # Si ninguno coincide exactamente, busca cualquiera que diga "SQL Server"
    for d in installed:
        if "sql server" in d.lower():
            return d
    return None

# ──────────────────────────────────────────────
# COLORES / TEMA
# ──────────────────────────────────────────────
C = {
    "bg":        "#0F1117",
    "surface":   "#1A1D27",
    "card":      "#22263A",
    "border":    "#2E3250",
    "accent":    "#3182DF",
    "accent2":   "#5BA3FF",
    "success":   "#21B868",
    "danger":    "#E74C3C",
    "warning":   "#F4A62A",
    "text":      "#E8ECF0",
    "text_dim":  "#7B8499",
    "text_mute": "#454B63",
    "white":     "#FFFFFF",
    "step_done": "#1A3A22",
    "step_act":  "#0D2340",
    "step_idle": "#181B2A",
    "log_bg":    "#0A0D14",
    "log_text":  "#A8F0C6",
    "log_err":   "#FF6B6B",
    "log_warn":  "#F4A62A",
    "log_info":  "#5BA3FF",
}

STEPS = ["1  Servidor", "2  Empresa", "3  Departamentos", "4  Exportar"]


# ──────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {}


def save_config(data):
    with open(CONFIG_FILE, "w") as f:
        json.dump(data, f, indent=4)


# ──────────────────────────────────────────────
# DETECCION DE SERVIDORES
# ──────────────────────────────────────────────
def _servers_from_net_use():
    servers = set()
    try:
        result = subprocess.run(
            ["net", "use"],
            capture_output=True, text=True, timeout=6,
            encoding="cp850", errors="replace",
        )
        for match in re.finditer(r"\\\\([^\\]+)\\", result.stdout):
            host = match.group(1).strip()
            if host:
                servers.add(host)
                servers.add(f"{host}\\SQLEXPRESS")
                servers.add(f"{host}\\MSSQLSERVER")
                servers.add(f"{host}\\COMPAC")
                servers.add(f"{host}\\NOMINAS")
    except Exception:
        pass
    return servers


def _servers_from_sqlcmd():
    found = set()
    for cmd in (["sqlcmd", "-L"], ["osql", "-L"]):
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=8)
            for line in result.stdout.splitlines():
                s = line.strip()
                if s and not s.lower().startswith("server"):
                    found.add(s)
        except Exception:
            pass
    return found


def _servers_from_registry():
    found = set()
    try:
        import winreg
        reg_path = r"SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
            i = 0
            while True:
                try:
                    instance_name, _, _ = winreg.EnumValue(key, i)
                    if instance_name.upper() == "MSSQLSERVER":
                        found.add("(local)")
                    else:
                        found.add(f"(local)\\{instance_name}")
                    i += 1
                except OSError:
                    break
    except Exception:
        pass
    return found


def detect_sql_servers():
    found = set()
    found |= _servers_from_sqlcmd()
    found |= _servers_from_registry()
    found |= _servers_from_net_use()
    for d in ["(local)", "localhost", r"localhost\SQLEXPRESS", r".\SQLEXPRESS"]:
        found.add(d)

    def sort_key(s):
        sl = s.lower()
        if sl.startswith("(local") or sl.startswith("local") or sl.startswith("."):
            return (1, sl)
        return (0, sl)
    return sorted(found, key=sort_key)


# ──────────────────────────────────────────────
# BASE DE DATOS
# ──────────────────────────────────────────────
def _make_conn_str(server, database=None, timeout=10, driver=None):
    if driver is None:
        driver = get_available_odbc_driver()
    if driver is None:
        raise RuntimeError(
            "No se encontro ningun driver ODBC de SQL Server instalado en este equipo.\n\n"
            "Instala 'ODBC Driver 17 for SQL Server' desde:\n"
            "https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server\n\n"
            "Luego REINICIA la computadora e intenta de nuevo."
        )
    base = (
        f"DRIVER={{{driver}}};"
        f"SERVER={server};Trusted_Connection=yes;"
        f"Connect Timeout={timeout};"
    )
    if database:
        base += f"DATABASE={database};"
    return base, driver


def test_connection(server):
    """
    Intenta conectar con cada driver disponible hasta encontrar uno que funcione.
    Retorna (True, driver_usado, None) o (False, None, mensaje_error).
    """
    installed = [d for d in pyodbc.drivers() if "sql server" in d.lower()]
    if not installed:
        return False, None, (
            "No hay ningun driver ODBC de SQL Server instalado.\n"
            "Instala 'ODBC Driver 17 for SQL Server' y reinicia."
        )

    last_error = ""
    for driver in ODBC_DRIVERS_PRIORITY:
        if driver not in installed:
            continue
        try:
            cs = (
                f"DRIVER={{{driver}}};"
                f"SERVER={server};Trusted_Connection=yes;"
                "Connect Timeout=5;"
            )
            conn = pyodbc.connect(cs)
            conn.close()
            return True, driver, None
        except Exception as e:
            last_error = str(e)

    # Si ninguno del priority list funciona, intenta con los instalados restantes
    for driver in installed:
        if driver in ODBC_DRIVERS_PRIORITY:
            continue
        try:
            cs = (
                f"DRIVER={{{driver}}};"
                f"SERVER={server};Trusted_Connection=yes;"
                "Connect Timeout=5;"
            )
            conn = pyodbc.connect(cs)
            conn.close()
            return True, driver, None
        except Exception as e:
            last_error = str(e)

    return False, None, last_error


def get_databases(server, driver):
    cs = f"DRIVER={{{driver}}};SERVER={server};Trusted_Connection=yes;Connect Timeout=10;"
    conn = pyodbc.connect(cs)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sys.databases WHERE name LIKE 'NOMINAS%' ORDER BY name")
    rows = [row[0] for row in cursor.fetchall()]
    conn.close()
    return rows


def get_departments(server, database, driver):
    cs = (f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};"
          "Trusted_Connection=yes;Connect Timeout=10;")
    conn = pyodbc.connect(cs)
    df = pd.read_sql(
        "SELECT cIdDepartamento, cNombreDepartamento "
        "FROM NomDepartamento ORDER BY cNombreDepartamento", conn)
    conn.close()
    return df


# ──────────────────────────────────────────────
# QUERY / EXPORT
# ──────────────────────────────────────────────
def build_query(only_active=True, selected_departments=None):
    where = []
    if only_active:
        where.append("E.cEstatus = 1")
    if selected_departments:
        ids = ",".join(map(str, selected_departments))
        where.append(f"E.cIdDepartamento IN ({ids})")
    where_clause = ("WHERE " + " AND ".join(where)) if where else ""
    return f"""
    SELECT
        E.cCodigoEmpleado            AS Codigo,
        E.cRFC                       AS RFC,
        E.cCURP                      AS CURP,
        E.cNombre                    AS Nombre,
        E.cSueldoDiario              AS [Salario Diario],
        E.cSalarioDiarioIntegrado    AS SDI,
        D.cNombreDepartamento        AS Departamento
    FROM NomEmpleado E
    LEFT JOIN NomDepartamento D ON E.cIdDepartamento = D.cIdDepartamento
    {where_clause}
    ORDER BY E.cCodigoEmpleado
    """


def export_data(server, database, output_path, only_active, selected_departments, driver):
    cs = (f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};"
          "Trusted_Connection=yes;Connect Timeout=20;")
    conn = pyodbc.connect(cs)
    df = pd.read_sql(build_query(only_active, selected_departments), conn)
    conn.close()

    if df.empty:
        raise ValueError("La consulta no arrojo resultados. Verifica los filtros.")

    df.to_excel(output_path, index=False, engine="openpyxl")
    wb = load_workbook(output_path)
    ws = wb.active
    ws.title = "Empleados"

    BLUE     = "3182DF"
    ROW_ALT  = "EEF4FB"
    BORDER_C = "BDD0E8"
    thin_b   = Border(**{s: Side(style="thin", color=BORDER_C)
                         for s in ("left", "right", "top", "bottom")})

    for cell in ws[1]:
        cell.fill      = PatternFill(start_color=BLUE, end_color=BLUE, fill_type="solid")
        cell.font      = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
        cell.border    = thin_b
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    max_row = ws.max_row
    for r in range(2, max_row + 1):
        fc = ROW_ALT if r % 2 == 0 else "FFFFFF"
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill      = PatternFill(start_color=fc, end_color=fc, fill_type="solid")
            cell.border    = thin_b
            cell.font      = Font(name="Calibri", size=10)
            cell.alignment = Alignment(vertical="center")

    for col in ws.columns:
        ml = max((len(str(cell.value)) for cell in col if cell.value), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(ml + 4, 50)

    tr = max_row + 2
    for col_idx, (val, align, color) in enumerate([
        ("TOTAL EMPLEADOS:", "right",  "0F1117"),
        (max_row - 1,        "center", BLUE),
    ], 1):
        cell = ws.cell(row=tr, column=col_idx, value=val)
        cell.font      = Font(bold=True, color=color, name="Calibri", size=11)
        cell.fill      = PatternFill(start_color="D6E4F5", end_color="D6E4F5", fill_type="solid")
        cell.alignment = Alignment(horizontal=align)

    ws.freeze_panes = "A2"
    wb.save(output_path)
    return len(df)


# ──────────────────────────────────────────────
# GUI
# ──────────────────────────────────────────────
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel SolutionsV \u00b7 Exportador Nominas ContpaqI")
        self.root.configure(bg=C["bg"])
        self.root.resizable(True, True)
        self.root.minsize(780, 780)

        self.config         = load_config()
        self._dept_ids      = []
        self._scanning      = False
        self._filter_active = True
        self._active_driver = None   # driver ODBC que funciono

        self._apply_styles()
        self._build_ui()
        self._center_window(820, 920)

        # Log inicial: drivers disponibles
        installed = [d for d in pyodbc.drivers() if "sql server" in d.lower()]
        if installed:
            self._log(f"Drivers ODBC encontrados: {', '.join(installed)}", "info")
        else:
            self._log(
                "ADVERTENCIA: No se encontro ningun driver ODBC de SQL Server.\n"
                "  -> Instala 'ODBC Driver 17 for SQL Server' y REINICIA el equipo.",
                "error")

        self.root.after(300, self._start_scan)

    # ── ttk styles ────────────────────────────────
    def _apply_styles(self):
        s = ttk.Style()
        s.theme_use("clam")
        s.configure("TFrame",       background=C["bg"])
        s.configure("TLabelframe",
                    background=C["card"], foreground=C["text_dim"],
                    bordercolor=C["border"], relief="flat", padding=10)
        s.configure("TLabelframe.Label",
                    background=C["card"], foreground=C["accent2"],
                    font=("Segoe UI", 9, "bold"))
        s.configure("TCombobox",
                    fieldbackground=C["surface"], background=C["surface"],
                    foreground=C["text"], selectbackground=C["accent"],
                    selectforeground=C["white"], bordercolor=C["border"],
                    arrowcolor=C["accent2"])
        s.map("TCombobox",
              fieldbackground=[("readonly", C["surface"])],
              foreground=[("readonly", C["text"])])
        s.configure("TEntry",
                    fieldbackground=C["surface"], foreground=C["text"],
                    bordercolor=C["border"], insertcolor=C["text"])
        s.configure("Vertical.TScrollbar",
                    background=C["surface"], troughcolor=C["bg"],
                    arrowcolor=C["text_dim"], bordercolor=C["border"])
        s.configure("TProgressbar",
                    troughcolor=C["surface"], background=C["success"],
                    bordercolor=C["border"])

    # ── UI ───────────────────────────────────────
    def _build_ui(self):
        root = self.root

        # Header
        hdr = tk.Frame(root, bg=C["accent"], height=52)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="  \u2b21  Excel SolutionsV",
                 bg=C["accent"], fg=C["white"],
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=18, pady=10)
        tk.Label(hdr, text="Exportador de Nominas ContpaqI",
                 bg=C["accent"], fg="#BFD9FF",
                 font=("Segoe UI", 10)).pack(side="left", pady=10)

        # Pipeline stepper
        self._step_labels = []
        step_bar = tk.Frame(root, bg=C["surface"], height=46)
        step_bar.pack(fill="x")
        step_bar.pack_propagate(False)
        inner = tk.Frame(step_bar, bg=C["surface"])
        inner.pack(fill="both", expand=True)
        for i, label in enumerate(STEPS):
            if i > 0:
                tk.Label(inner, text=" \u203a ", bg=C["surface"], fg=C["text_mute"],
                         font=("Segoe UI", 14, "bold")).pack(side="left")
            lbl = tk.Label(inner, text=f"  {label}  ",
                           bg=C["step_idle"], fg=C["text_mute"],
                           font=("Segoe UI", 9, "bold"),
                           padx=4, pady=5, relief="flat")
            lbl.pack(side="left", fill="y", padx=2, pady=5)
            self._step_labels.append(lbl)
        self._set_step(0)

        # Scroll container para el cuerpo principal
        canvas_frame = tk.Frame(root, bg=C["bg"])
        canvas_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(canvas_frame, bg=C["bg"], highlightthickness=0)
        vscroll = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        vscroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        body = tk.Frame(canvas, bg=C["bg"], padx=16, pady=14)
        body_id = canvas.create_window((0, 0), window=body, anchor="nw")

        def _on_resize(event):
            canvas.itemconfig(body_id, width=event.width)
        canvas.bind("<Configure>", _on_resize)

        def _on_body_resize(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        body.bind("<Configure>", _on_body_resize)

        # Mousewheel
        def _on_wheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_wheel)

        # ── PASO 1: Servidor ──
        s1 = ttk.LabelFrame(body, text="  \U0001f50c  Paso 1 \u2014 Servidor SQL")
        s1.pack(fill="x", pady=(0, 8))

        tk.Label(s1, text="Servidor:", bg=C["card"], fg=C["text_dim"],
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.server_combo = ttk.Combobox(s1, width=36, state="normal", font=("Segoe UI", 10))
        self.server_combo.grid(row=0, column=1, sticky="ew", pady=4)
        self.server_combo.bind("<<ComboboxSelected>>", lambda e: self._on_server_pick())
        saved = self.config.get("server", "")
        if saved:
            self.server_combo.set(saved)

        self.scan_btn = self._mk_btn(s1, "\U0001f50d Buscar en red", self._start_scan, "secondary")
        self.scan_btn.grid(row=0, column=2, padx=(8, 0))

        self.conn_indicator = tk.Label(
            s1, text="\u25cf  Sin probar",
            bg=C["card"], fg=C["text_mute"],
            font=("Segoe UI", 9, "bold"))
        self.conn_indicator.grid(row=1, column=1, sticky="w", pady=(2, 4))

        self._mk_btn(s1, "\u26a1 Probar conexion", self._test_conn, "ghost").grid(
            row=1, column=2, padx=(8, 0), pady=(2, 4))
        s1.columnconfigure(1, weight=1)

        # ── PASO 2: Empresa ──
        s2 = ttk.LabelFrame(body, text="  \U0001f3e2  Paso 2 \u2014 Empresa (base de datos)")
        s2.pack(fill="x", pady=(0, 8))

        tk.Label(s2, text="Empresa:", bg=C["card"], fg=C["text_dim"],
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.db_combo = ttk.Combobox(s2, width=36, state="readonly", font=("Segoe UI", 10))
        self.db_combo.grid(row=0, column=1, sticky="ew", pady=4)
        self.db_combo.bind("<<ComboboxSelected>>", lambda e: self._load_departments())
        self.load_db_btn = self._mk_btn(s2, "\u21bb Cargar empresas", self._load_databases, "primary")
        self.load_db_btn.grid(row=0, column=2, padx=(8, 0))
        s2.columnconfigure(1, weight=1)

        # ── PASO 3: Departamentos ──
        s3 = ttk.LabelFrame(
            body,
            text="  \U0001f4c2  Paso 3 \u2014 Departamentos  "
                 "(Ctrl+clic = multiple  |  Sin seleccion = todos)")
        s3.pack(fill="x", pady=(0, 8))

        lw = tk.Frame(s3, bg=C["card"])
        lw.pack(fill="x")
        self.dept_listbox = tk.Listbox(
            lw, selectmode=tk.MULTIPLE,
            bg=C["surface"], fg=C["text"],
            selectbackground=C["accent"], selectforeground=C["white"],
            font=("Consolas", 10), height=5, relief="flat", bd=0,
            activestyle="none",
            highlightthickness=1,
            highlightcolor=C["border"], highlightbackground=C["border"])
        self.dept_listbox.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(lw, orient="vertical", command=self.dept_listbox.yview)
        sb.pack(side="right", fill="y")
        self.dept_listbox.configure(yscrollcommand=sb.set)

        # ── PASO 4: Filtro toggle ──
        s4 = ttk.LabelFrame(body, text="  \U0001f3af  Paso 4 \u2014 \u00bfQue empleados exportar?")
        s4.pack(fill="x", pady=(0, 8))

        toggle_row = tk.Frame(s4, bg=C["card"])
        toggle_row.pack(fill="x")

        self._tog_activos = self._mk_toggle(
            toggle_row, label="\u2714  Solo Activos",
            sublabel="Excluye bajas y suspendidos",
            selected=True, command=lambda: self._set_filter(True))
        self._tog_activos.pack(side="left", fill="both", expand=True, padx=(0, 6), pady=4)

        self._tog_todos = self._mk_toggle(
            toggle_row, label="\u229e  Todos los Empleados",
            sublabel="Incluye bajas, suspendidos e inactivos",
            selected=False, command=lambda: self._set_filter(False))
        self._tog_todos.pack(side="left", fill="both", expand=True, pady=4)

        # ── PASO 5: Archivo ──
        s5 = ttk.LabelFrame(body, text="  \U0001f4be  Paso 5 \u2014 Archivo de salida")
        s5.pack(fill="x", pady=(0, 10))

        tk.Label(s5, text="Ruta:", bg=C["card"], fg=C["text_dim"],
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.file_entry = ttk.Entry(s5, width=38, font=("Segoe UI", 10))
        self.file_entry.grid(row=0, column=1, sticky="ew", pady=4)
        last_out = self.config.get("last_output", "")
        if last_out:
            self.file_entry.insert(0, last_out)
        self._mk_btn(s5, "\U0001f4c1 Examinar", self._select_file, "ghost").grid(
            row=0, column=2, padx=(8, 0))
        s5.columnconfigure(1, weight=1)

        # ── Progress ──
        self.progress = ttk.Progressbar(body, mode="indeterminate")

        # ── Boton GENERAR (verde) ──
        self.gen_btn = tk.Button(
            body,
            text="  \u25b6   Generar Reporte Excel",
            command=self._execute,
            bg=C["success"], fg=C["white"],
            activebackground="#189B52", activeforeground=C["white"],
            font=("Segoe UI", 12, "bold"),
            relief="flat", cursor="hand2", pady=14,
        )
        self.gen_btn.pack(fill="x", pady=(0, 10))

        # ── LOG ──
        log_header = tk.Frame(body, bg=C["bg"])
        log_header.pack(fill="x", pady=(0, 4))

        tk.Label(log_header, text="\U0001f4cb  Log de actividad",
                 bg=C["bg"], fg=C["text_dim"],
                 font=("Segoe UI", 9, "bold")).pack(side="left")

        self._mk_btn(log_header, "\U0001f4cb Copiar log", self._copy_log, "ghost").pack(
            side="right")
        self._mk_btn(log_header, "\u2715 Limpiar", self._clear_log, "ghost").pack(
            side="right", padx=(0, 6))

        self.log_box = tk.Text(
            body,
            bg=C["log_bg"], fg=C["log_text"],
            font=("Consolas", 9),
            height=10, relief="flat", bd=0,
            state="disabled",
            highlightthickness=1,
            highlightcolor=C["border"], highlightbackground=C["border"],
            wrap="word",
            insertbackground=C["log_text"],
        )
        self.log_box.pack(fill="x", pady=(0, 6))

        # Tags de color para el log
        self.log_box.tag_config("info",    foreground=C["log_info"])
        self.log_box.tag_config("success", foreground=C["success"])
        self.log_box.tag_config("warning", foreground=C["warning"])
        self.log_box.tag_config("error",   foreground=C["log_err"])
        self.log_box.tag_config("default", foreground=C["log_text"])
        self.log_box.tag_config("ts",      foreground=C["text_mute"])

        # ── Barra de estado ──
        status_bar = tk.Frame(root, bg=C["surface"], height=28)
        status_bar.pack(fill="x", side="bottom")
        status_bar.pack_propagate(False)
        self._status_var = tk.StringVar(value="Listo.")
        tk.Label(status_bar, textvariable=self._status_var,
                 bg=C["surface"], fg=C["text_dim"],
                 font=("Segoe UI", 9), anchor="w").pack(side="left", padx=10)
        tk.Label(status_bar, text="Excel SolutionsV \u00a9 2025",
                 bg=C["surface"], fg=C["text_mute"],
                 font=("Segoe UI", 8)).pack(side="right", padx=10)

    # ── Log ──────────────────────────────────────
    def _log(self, msg, level="default"):
        """Agrega una linea al log con timestamp y color segun nivel."""
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        prefix = {"info": "[INFO]", "success": "[OK]  ",
                  "warning": "[WARN]", "error": "[ERR] "}.get(level, "[LOG] ")

        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{ts}  ", "ts")
        self.log_box.insert("end", f"{prefix}  ", level)
        self.log_box.insert("end", msg + "\n", level)
        self.log_box.configure(state="disabled")
        self.log_box.see("end")

    def _copy_log(self):
        content = self.log_box.get("1.0", "end")
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self._set_status("Log copiado al portapapeles.")

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    # ── Widget helpers ────────────────────────────
    def _mk_btn(self, parent, text, command, kind="primary"):
        cfg = {
            "primary":   (C["accent"],  C["white"],    "#255EAA"),
            "secondary": (C["surface"], C["accent2"],  C["border"]),
            "ghost":     (C["card"],    C["text_dim"], C["border"]),
        }
        bg, fg, abg = cfg[kind]
        return tk.Button(parent, text=text, command=command,
                         bg=bg, fg=fg,
                         activebackground=abg, activeforeground=C["white"],
                         font=("Segoe UI", 9, "bold" if kind == "primary" else "normal"),
                         relief="flat", cursor="hand2", padx=10, pady=5)

    def _mk_toggle(self, parent, label, sublabel, selected, command):
        border_color = C["accent"] if selected else C["border"]
        frame = tk.Frame(parent, bg=C["card"], cursor="hand2",
                         highlightthickness=2,
                         highlightbackground=border_color)
        title_fg = C["accent2"] if selected else C["text_dim"]
        sub_fg   = C["text_dim"] if selected else C["text_mute"]
        title = tk.Label(frame, text=label, bg=C["card"], fg=title_fg,
                         font=("Segoe UI", 11, "bold"), anchor="w", padx=10)
        title.pack(fill="x", pady=(8, 1))
        sub = tk.Label(frame, text=sublabel, bg=C["card"], fg=sub_fg,
                       font=("Segoe UI", 8), anchor="w", padx=12)
        sub.pack(fill="x", pady=(0, 8))
        frame._title_lbl = title
        frame._sub_lbl   = sub
        for w in (frame, title, sub):
            w.bind("<Button-1>", lambda e: command())
        return frame

    def _set_filter(self, only_active: bool):
        self._filter_active = only_active
        for frame, is_act in [(self._tog_activos, True), (self._tog_todos, False)]:
            sel = (only_active == is_act)
            frame.configure(highlightbackground=C["accent"] if sel else C["border"])
            frame._title_lbl.configure(fg=C["accent2"] if sel else C["text_dim"])
            frame._sub_lbl.configure(fg=C["text_dim"] if sel else C["text_mute"])

    # ── Stepper ───────────────────────────────────
    def _set_step(self, step: int):
        for i, lbl in enumerate(self._step_labels):
            if i < step:
                lbl.configure(bg=C["step_done"], fg=C["success"],
                               text=f"  \u2714 {STEPS[i]}  ")
            elif i == step:
                lbl.configure(bg=C["step_act"], fg=C["accent2"],
                               text=f"  \u25ba {STEPS[i]}  ")
            else:
                lbl.configure(bg=C["step_idle"], fg=C["text_mute"],
                               text=f"  {STEPS[i]}  ")

    # ── Helpers ───────────────────────────────────
    def _center_window(self, w, h):
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _set_status(self, msg):
        self._status_var.set(msg)

    def _start_progress(self):
        self.progress.pack(fill="x", pady=(0, 6), before=self.gen_btn)
        self.progress.start(10)

    def _stop_progress(self):
        self.progress.stop()
        self.progress.pack_forget()

    # ── Scan servidores ───────────────────────────
    def _start_scan(self):
        if self._scanning:
            return
        self._scanning = True
        self.scan_btn.configure(text="\u23f3 Buscando...", state="disabled")
        self._set_status("Detectando servidores...")
        self._log("Iniciando deteccion de servidores SQL en la red...", "info")
        threading.Thread(
            target=lambda: self.root.after(0, self._scan_done, detect_sql_servers()),
            daemon=True).start()

    def _scan_done(self, servers):
        self._scanning = False
        current = self.server_combo.get()
        all_svr = list(servers)
        if current and current not in all_svr:
            all_svr.insert(0, current)
        self.server_combo["values"] = all_svr
        if not current and all_svr:
            self.server_combo.set(all_svr[0])
        self.scan_btn.configure(text="\U0001f50d Buscar en red", state="normal")
        self._log(f"Deteccion completada. {len(servers)} opciones disponibles.", "success")
        for s in servers:
            self._log(f"  -> {s}", "default")
        self._set_status(f"Deteccion completada — {len(servers)} opciones.")

    def _on_server_pick(self):
        self._set_step(0)
        self._active_driver = None
        self.conn_indicator.configure(text="\u25cf  Sin probar", fg=C["text_mute"])
        self._log(f"Servidor seleccionado: {self.server_combo.get()}", "info")

    # ── Probar conexion ───────────────────────────
    def _test_conn(self):
        server = self.server_combo.get().strip()
        if not server:
            messagebox.showwarning("Aviso", "Selecciona el servidor primero.")
            return
        self.conn_indicator.configure(text="\u25cf  Probando...", fg=C["warning"])
        self._set_status(f"Probando conexion a {server}...")
        self._log(f"Probando conexion a: {server}", "info")

        # Log de drivers disponibles
        installed = [d for d in pyodbc.drivers() if "sql server" in d.lower()]
        self._log(f"Drivers ODBC disponibles: {installed if installed else 'NINGUNO'}", "info")

        threading.Thread(
            target=lambda: self.root.after(0, self._test_done, server, *test_connection(server)),
            daemon=True).start()

    def _test_done(self, server, ok, driver, err):
        if ok:
            self._active_driver = driver
            self.conn_indicator.configure(text=f"\u25cf  Conectado \u2714  [{driver}]",
                                          fg=C["success"])
            self._log(f"Conexion exitosa usando driver: {driver}", "success")
            self._set_status(f"Conectado con driver: {driver}")
            self._set_step(1)
        else:
            self.conn_indicator.configure(text="\u25cf  Fallo la conexion", fg=C["danger"])
            self._log(f"Error al conectar: {err}", "error")
            self._set_status("Fallo la conexion. Revisa el log para detalles.")

    # ── Cargar empresas ───────────────────────────
    def _load_databases(self):
        server = self.server_combo.get().strip()
        if not server:
            messagebox.showwarning("Aviso", "Selecciona el servidor primero.")
            return
        if not self._active_driver:
            messagebox.showwarning("Aviso",
                                   "Primero prueba la conexion con '⚡ Probar conexion'.")
            return
        self._set_status("Cargando empresas...")
        self._log(f"Buscando bases de datos NOMINAS en {server}...", "info")
        self._start_progress()
        self.load_db_btn.configure(state="disabled")

        def _work():
            try:
                dbs = get_databases(server, self._active_driver)
                self.root.after(0, self._dbs_loaded, dbs)
            except Exception as e:
                self.root.after(0, self._err_in_thread, str(e))

        threading.Thread(target=_work, daemon=True).start()

    def _dbs_loaded(self, dbs):
        self._stop_progress()
        self.load_db_btn.configure(state="normal")
        if not dbs:
            self._log("No se encontraron bases de datos NOMINAS* en este servidor.", "warning")
            messagebox.showwarning("Sin resultados",
                                   "No se encontraron bases de datos NOMINAS* en este servidor.")
            return
        self.db_combo["values"] = dbs
        self.db_combo.current(0)
        self._log(f"Empresas encontradas: {', '.join(dbs)}", "success")
        self._set_status(f"{len(dbs)} empresa(s) encontrada(s).")
        self._set_step(2)
        self._load_departments()

    # ── Cargar departamentos ──────────────────────
    def _load_departments(self):
        server = self.server_combo.get().strip()
        db     = self.db_combo.get().strip()
        if not server or not db or not self._active_driver:
            return
        self._log(f"Cargando departamentos de '{db}'...", "info")
        self.dept_listbox.delete(0, tk.END)
        self._dept_ids = []

        def _work():
            try:
                df = get_departments(server, db, self._active_driver)
                self.root.after(0, self._depts_loaded, df)
            except Exception as e:
                self.root.after(0, self._err_in_thread, str(e))

        threading.Thread(target=_work, daemon=True).start()

    def _depts_loaded(self, df):
        for _, row in df.iterrows():
            self._dept_ids.append(row["cIdDepartamento"])
            self.dept_listbox.insert(
                tk.END,
                f"  {str(row['cIdDepartamento']):>4}  \u2502  {row['cNombreDepartamento']}")
        self._log(f"{len(df)} departamento(s) cargado(s).", "success")
        self._set_status(f"{len(df)} departamento(s) listos.")
        self._set_step(3)

    # ── Archivo ───────────────────────────────────
    def _select_file(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
            title="Guardar reporte como...")
        if path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, path)
            self._log(f"Archivo de salida: {path}", "info")

    # ── Exportar ─────────────────────────────────
    def _execute(self):
        server    = self.server_combo.get().strip()
        db        = self.db_combo.get().strip()
        file_path = self.file_entry.get().strip()

        missing = []
        if not server:              missing.append("  \u2022 Servidor SQL")
        if not db:                  missing.append("  \u2022 Empresa")
        if not self._active_driver: missing.append("  \u2022 Conexion probada (usa Probar conexion)")
        if not file_path:           missing.append("  \u2022 Archivo de salida")
        if missing:
            messagebox.showwarning("Faltan datos",
                                   "Completa:\n\n" + "\n".join(missing))
            return

        only_active = self._filter_active
        sel_idx     = self.dept_listbox.curselection()
        sel_ids     = [self._dept_ids[i] for i in sel_idx] if sel_idx else None

        filtro_str = "Solo activos" if only_active else "Todos"
        depts_str  = f"{len(sel_idx)} departamento(s)" if sel_idx else "todos los departamentos"
        self._log(f"Iniciando exportacion | Filtro: {filtro_str} | Depts: {depts_str}", "info")
        self._log(f"  Servidor: {server}  |  Empresa: {db}", "info")
        self._log(f"  Driver:   {self._active_driver}", "info")
        self._log(f"  Destino:  {file_path}", "info")

        self.gen_btn.configure(state="disabled", text="  \u23f3  Generando...")
        self._start_progress()
        self._set_status("Exportando datos...")

        def _work():
            try:
                n = export_data(server, db, file_path, only_active,
                                sel_ids, self._active_driver)
                self.root.after(0, self._export_done, n, file_path, server)
            except Exception as e:
                self.root.after(0, self._err_in_thread, str(e))

        threading.Thread(target=_work, daemon=True).start()

    def _export_done(self, n_rows, file_path, server):
        self._stop_progress()
        self.gen_btn.configure(state="normal", text="  \u25b6   Generar Reporte Excel")
        for i, lbl in enumerate(self._step_labels):
            lbl.configure(bg=C["step_done"], fg=C["success"],
                          text=f"  \u2714 {STEPS[i]}  ")
        self._log(f"Exportacion exitosa: {n_rows} empleado(s) -> {file_path}", "success")
        self._set_status(f"\u2714 {n_rows} empleado(s) exportados -> {file_path}")
        save_config({"server": server, "last_output": file_path})
        messagebox.showinfo("Listo \u2705",
                            f"Exportacion exitosa\n\nEmpleados: {n_rows}\nArchivo: {file_path}")
        if messagebox.askyesno("\u00bfAbrir?", "\u00bfAbrir el Excel ahora?"):
            os.startfile(file_path)

    def _err_in_thread(self, msg):
        self._stop_progress()
        self.gen_btn.configure(state="normal", text="  \u25b6   Generar Reporte Excel")
        self._log(f"ERROR: {msg}", "error")
        self._set_status(f"\u2718 Error. Revisa el log.")
        messagebox.showerror("Error", f"Ocurrio un error:\n\n{msg}")


# ──────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()