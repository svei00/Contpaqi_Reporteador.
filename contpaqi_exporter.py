# ==========================
# contpaqi_export_pro.py (ULTRA PRO v6)
# Excel SolutionsV — Exportador de Nominas ContpaqI
# ==========================

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
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
# El driver generico "SQL Server" no soporta Connect Timeout en la cadena,
# por eso usamos pyodbc.connect(..., timeout=N) en su lugar.
# ──────────────────────────────────────────────
ODBC_DRIVERS_PRIORITY = [
    "ODBC Driver 17 for SQL Server",
    "ODBC Driver 18 for SQL Server",
    "ODBC Driver 13 for SQL Server",
    "SQL Server Native Client 11.0",
    "SQL Server Native Client 10.0",
    "SQL Server",
]

def get_installed_sql_drivers():
    return [d for d in pyodbc.drivers() if "sql server" in d.lower()]

def pick_best_driver():
    installed = get_installed_sql_drivers()
    for d in ODBC_DRIVERS_PRIORITY:
        if d in installed:
            return d
    return installed[0] if installed else None


# ──────────────────────────────────────────────
# COLORES
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
def _net_use_hosts():
    hosts = set()
    try:
        r = subprocess.run(["net", "use"], capture_output=True, text=True,
                           timeout=6, encoding="cp850", errors="replace")
        for m in re.finditer(r"\\\\([^\\]+)\\", r.stdout):
            h = m.group(1).strip()
            if h:
                hosts.add(h)
    except Exception:
        pass
    return hosts

def _sqlcmd_servers():
    found = set()
    for cmd in (["sqlcmd", "-L"], ["osql", "-L"]):
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=8)
            for line in r.stdout.splitlines():
                s = line.strip()
                if s and not s.lower().startswith("server"):
                    found.add(s)
        except Exception:
            pass
    return found

def _registry_instances():
    found = set()
    try:
        import winreg
        path = r"SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
            i = 0
            while True:
                try:
                    name, _, _ = winreg.EnumValue(key, i)
                    found.add("(local)" if name.upper() == "MSSQLSERVER"
                              else f"(local)\\{name}")
                    i += 1
                except OSError:
                    break
    except Exception:
        pass
    return found

def detect_sql_servers():
    found = set()
    found |= _sqlcmd_servers()
    found |= _registry_instances()

    for host in _net_use_hosts():
        # ContpaqI SIEMPRE usa instancia \COMPAC — va primero
        found.add(f"{host}\\COMPAC")
        found.add(host)
        found.add(f"{host}\\SQLEXPRESS")
        found.add(f"{host}\\MSSQLSERVER")

    for d in ["(local)\\COMPAC", "(local)", "localhost\\COMPAC",
              "localhost", r".\COMPAC", r".\SQLEXPRESS"]:
        found.add(d)

    def key(s):
        sl = s.lower()
        # Prioridad: hosts de red con \COMPAC primero
        if "compac" in sl and not sl.startswith("(") and not sl.startswith(".") and not sl.startswith("local"):
            return (0, sl)
        if not sl.startswith("(") and not sl.startswith(".") and not sl.startswith("local"):
            return (1, sl)
        return (2, sl)

    return sorted(found, key=key)


# ──────────────────────────────────────────────
# CONEXION
# ──────────────────────────────────────────────
def _build_cs(server, driver, user=None, password=None, database=None):
    """
    Construye la cadena de conexion.
    - Sin user/password: Windows auth (Trusted_Connection).
      Solo funciona si cliente y servidor estan en el mismo dominio.
    - Con user/password: SQL Server auth. Necesario para conexiones remotas
      (ej. via Tailscale entre diferentes redes).
    NOTA: NO incluimos Connect Timeout aqui porque el driver generico
    'SQL Server' no lo soporta como atributo de cadena. El timeout
    se pasa directamente a pyodbc.connect().
    """
    if user and password:
        auth = f"UID={user};PWD={password};"
    else:
        auth = "Trusted_Connection=yes;"

    cs = f"DRIVER={{{driver}}};SERVER={server};{auth}"
    if database:
        cs += f"DATABASE={database};"
    return cs

def test_connection(server, user=None, password=None, timeout=8):
    """
    Intenta conectar probando cada driver instalado.
    Retorna (True, driver_usado, None) o (False, None, ultimo_error).
    """
    installed = get_installed_sql_drivers()
    if not installed:
        return False, None, (
            "No hay ningun driver ODBC de SQL Server instalado.\n"
            "Instala 'ODBC Driver 17 for SQL Server' y REINICIA el equipo.\n"
            "https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server"
        )

    last_err = ""
    # Ordena segun prioridad
    to_try = [d for d in ODBC_DRIVERS_PRIORITY if d in installed]
    to_try += [d for d in installed if d not in ODBC_DRIVERS_PRIORITY]

    for driver in to_try:
        try:
            cs = _build_cs(server, driver, user, password)
            conn = pyodbc.connect(cs, timeout=timeout)
            conn.close()
            return True, driver, None
        except Exception as e:
            last_err = str(e)

    return False, None, last_err

def get_databases(server, driver, user=None, password=None):
    cs = _build_cs(server, driver, user, password)
    conn = pyodbc.connect(cs, timeout=15)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sys.databases WHERE name LIKE 'NOMINAS%' ORDER BY name")
    rows = [r[0] for r in cursor.fetchall()]
    conn.close()
    return rows

def get_departments(server, database, driver, user=None, password=None):
    cs = _build_cs(server, driver, user, password, database)
    conn = pyodbc.connect(cs, timeout=15)
    df = pd.read_sql(
        "SELECT cIdDepartamento, cNombreDepartamento "
        "FROM NomDepartamento ORDER BY cNombreDepartamento", conn)
    conn.close()
    return df

def export_data(server, database, output_path, only_active,
                selected_departments, driver, user=None, password=None):
    cs = _build_cs(server, driver, user, password, database)
    conn = pyodbc.connect(cs, timeout=30)

    where = []
    if only_active:
        where.append("E.cEstatus = 1")
    if selected_departments:
        ids = ",".join(map(str, selected_departments))
        where.append(f"E.cIdDepartamento IN ({ids})")
    wc = ("WHERE " + " AND ".join(where)) if where else ""

    query = f"""
    SELECT
        E.cCodigoEmpleado         AS Codigo,
        E.cRFC                    AS RFC,
        E.cCURP                   AS CURP,
        E.cNombre                 AS Nombre,
        E.cSueldoDiario           AS [Salario Diario],
        E.cSalarioDiarioIntegrado AS SDI,
        D.cNombreDepartamento     AS Departamento
    FROM NomEmpleado E
    LEFT JOIN NomDepartamento D ON E.cIdDepartamento = D.cIdDepartamento
    {wc}
    ORDER BY E.cCodigoEmpleado
    """
    df = pd.read_sql(query, conn)
    conn.close()

    if df.empty:
        raise ValueError("La consulta no arrojo resultados. Verifica los filtros.")

    df.to_excel(output_path, index=False, engine="openpyxl")
    wb = load_workbook(output_path)
    ws = wb.active
    ws.title = "Empleados"

    BLUE = "3182DF"; ROW_ALT = "EEF4FB"; BC = "BDD0E8"
    thin = Border(**{s: Side(style="thin", color=BC)
                     for s in ("left", "right", "top", "bottom")})

    for cell in ws[1]:
        cell.fill      = PatternFill(start_color=BLUE, end_color=BLUE, fill_type="solid")
        cell.font      = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
        cell.border    = thin
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    mr = ws.max_row
    for r in range(2, mr + 1):
        fc = ROW_ALT if r % 2 == 0 else "FFFFFF"
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = PatternFill(start_color=fc, end_color=fc, fill_type="solid")
            cell.border = thin
            cell.font   = Font(name="Calibri", size=10)
            cell.alignment = Alignment(vertical="center")

    for col in ws.columns:
        ml = max((len(str(c.value)) for c in col if c.value), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(ml + 4, 50)

    tr = mr + 2
    for ci, (val, align, color) in enumerate([
        ("TOTAL EMPLEADOS:", "right", "0F1117"),
        (mr - 1, "center", BLUE)
    ], 1):
        cell = ws.cell(row=tr, column=ci, value=val)
        cell.font = Font(bold=True, color=color, name="Calibri", size=11)
        cell.fill = PatternFill(start_color="D6E4F5", end_color="D6E4F5", fill_type="solid")
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
        self.root.minsize(800, 820)

        self.config         = load_config()
        self._dept_ids      = []
        self._scanning      = False
        self._filter_active = True
        self._active_driver = None
        self._use_sql_auth  = tk.BooleanVar(value=True)   # default: SQL auth para conexiones remotas

        self._apply_styles()
        self._build_ui()
        self._center_window(860, 980)

        # Log inicial
        installed = get_installed_sql_drivers()
        if installed:
            self._log(f"Drivers ODBC encontrados: {', '.join(installed)}", "info")
            best = pick_best_driver()
            if "Driver 17" not in best and "Driver 18" not in best:
                self._log(
                    f"AVISO: Usando driver '{best}'. Para mejor compatibilidad instala "
                    "'ODBC Driver 17 for SQL Server' (ver README).", "warning")
        else:
            self._log(
                "ERROR CRITICO: No hay ningun driver ODBC de SQL Server instalado.\n"
                "  Instala 'ODBC Driver 17 for SQL Server' y REINICIA el equipo.\n"
                "  Consulta el README para instrucciones.", "error")

        self._log(
            "NOTA: Para conexiones remotas (Tailscale, VPN, red diferente) USA autenticacion SQL.\n"
            "  Windows auth (Trusted Connection) solo funciona en el mismo dominio.", "warning")

        self.root.after(400, self._start_scan)

    # ── Styles ───────────────────────────────────
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
                    arrowcolor=C["accent2"], insertcolor=C["white"])
        s.map("TCombobox",
              fieldbackground=[("readonly", C["surface"])],
              foreground=[("readonly", C["text"])],
              insertcolor=[("focus", C["white"])])
        s.configure("TEntry",
                    fieldbackground=C["surface"], foreground=C["text"],
                    bordercolor=C["border"], insertcolor=C["text"])
        s.configure("Vertical.TScrollbar",
                    background=C["surface"], troughcolor=C["bg"],
                    arrowcolor=C["text_dim"], bordercolor=C["border"])
        s.configure("TProgressbar",
                    troughcolor=C["surface"], background=C["success"],
                    bordercolor=C["border"])
        s.configure("TCheckbutton",
                    background=C["card"], foreground=C["text"],
                    font=("Segoe UI", 9))

    # ── Build UI ─────────────────────────────────
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
                           font=("Segoe UI", 9, "bold"), padx=4, pady=5)
            lbl.pack(side="left", fill="y", padx=2, pady=5)
            self._step_labels.append(lbl)
        self._set_step(0)

        # Scrollable body
        wrap = tk.Frame(root, bg=C["bg"])
        wrap.pack(fill="both", expand=True)
        canvas = tk.Canvas(wrap, bg=C["bg"], highlightthickness=0)
        vscroll = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        vscroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        body = tk.Frame(canvas, bg=C["bg"], padx=16, pady=14)
        bid = canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(bid, width=e.width))
        body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        # ── PASO 1: Servidor ──────────────────────
        s1 = ttk.LabelFrame(body, text="  \U0001f50c  Paso 1 \u2014 Servidor SQL")
        s1.pack(fill="x", pady=(0, 8))

        tk.Label(s1, text="Servidor:", bg=C["card"], fg=C["text_dim"],
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        self.server_combo = ttk.Combobox(s1, width=34, state="normal", font=("Segoe UI", 10))
        self.server_combo.grid(row=0, column=1, sticky="ew", pady=3)
        self.server_combo.bind("<<ComboboxSelected>>", lambda e: self._on_server_pick())
        if self.config.get("server"):
            self.server_combo.set(self.config["server"])

        self.scan_btn = self._mk_btn(s1, "\U0001f50d Buscar", self._start_scan, "secondary")
        self.scan_btn.grid(row=0, column=2, padx=(6, 0), pady=3)

        # ── Autenticacion ─────────────────────────
        auth_frame = tk.Frame(s1, bg=C["card"])
        auth_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(6, 2))

        # Toggle Windows / SQL auth
        auth_label = tk.Label(auth_frame, text="Autenticacion:",
                              bg=C["card"], fg=C["text_dim"], font=("Segoe UI", 9))
        auth_label.pack(side="left", padx=(0, 10))

        self._auth_win_btn = tk.Button(
            auth_frame, text="Windows (mismo dominio)",
            command=lambda: self._set_auth_mode(False),
            bg=C["surface"], fg=C["text_mute"],
            activebackground=C["border"], activeforeground=C["white"],
            font=("Segoe UI", 9), relief="flat", cursor="hand2", padx=8, pady=3)
        self._auth_win_btn.pack(side="left", padx=(0, 4))

        self._auth_sql_btn = tk.Button(
            auth_frame, text="SQL Server (usuario/password)",
            command=lambda: self._set_auth_mode(True),
            bg=C["accent"], fg=C["white"],
            activebackground="#255EAA", activeforeground=C["white"],
            font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2", padx=8, pady=3)
        self._auth_sql_btn.pack(side="left")

        tk.Label(auth_frame, text="  <- Usa esto para conexiones remotas/VPN",
                 bg=C["card"], fg=C["warning"], font=("Segoe UI", 8, "italic")).pack(
            side="left", padx=(8, 0))

        # Campos SQL auth
        self._sql_auth_frame = tk.Frame(s1, bg=C["card"])
        self._sql_auth_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(4, 2))

        tk.Label(self._sql_auth_frame, text="Usuario:",
                 bg=C["card"], fg=C["text_dim"], font=("Segoe UI", 9)).grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        self.user_entry = tk.Entry(
            self._sql_auth_frame, width=20, font=("Segoe UI", 10),
            bg=C["surface"], fg=C["text"], insertbackground=C["white"],
            relief="flat", highlightthickness=1,
            highlightcolor=C["accent"], highlightbackground=C["border"])
        self.user_entry.grid(row=0, column=1, sticky="w", pady=3)
        self.user_entry.insert(0, self.config.get("db_user", "sa"))

        tk.Label(self._sql_auth_frame, text="Usuario admin de SQL  (ContpaqI usa 'sa')",
                 bg=C["card"], fg=C["text_mute"], font=("Segoe UI", 8)).grid(
            row=0, column=2, sticky="w", padx=(8, 0))

        tk.Label(self._sql_auth_frame, text="Password:",
                 bg=C["card"], fg=C["text_dim"], font=("Segoe UI", 9)).grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=3)
        self.pass_entry = tk.Entry(
            self._sql_auth_frame, width=20, font=("Segoe UI", 10),
            bg=C["surface"], fg=C["text"], insertbackground=C["white"],
            show="\u2022", relief="flat", highlightthickness=1,
            highlightcolor=C["accent"], highlightbackground=C["border"])
        self.pass_entry.grid(row=1, column=1, sticky="w", pady=3)

        # Controles: mostrar password + sugerencias de passwords comunes
        pass_ctrl = tk.Frame(self._sql_auth_frame, bg=C["card"])
        pass_ctrl.grid(row=1, column=2, sticky="w", padx=(8, 0))

        self._show_pass_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            pass_ctrl, text="Mostrar",
            variable=self._show_pass_var,
            command=self._toggle_pass_visibility,
            bg=C["card"], fg=C["text_dim"],
            activebackground=C["card"], selectcolor=C["surface"],
            font=("Segoe UI", 8), relief="flat"
        ).pack(side="left")

        # Dropdown con passwords comunes de ContpaqI
        COMMON_PASSWORDS = ["contpaqi", "Contpaqi1", "(en blanco)", "sa", "adminsa", "123456"]
        pwd_hint_btn = tk.Menubutton(
            pass_ctrl, text="\u25be Probar comunes",
            bg=C["surface"], fg=C["accent2"],
            font=("Segoe UI", 8), relief="flat", cursor="hand2",
            activebackground=C["border"], activeforeground=C["white"])
        pwd_hint_btn.pack(side="left", padx=(6, 0))
        pwd_menu = tk.Menu(pwd_hint_btn, tearoff=0,
                           bg=C["surface"], fg=C["text"],
                           activebackground=C["accent"], activeforeground=C["white"],
                           font=("Consolas", 9))
        pwd_hint_btn["menu"] = pwd_menu
        for pw in COMMON_PASSWORDS:
            actual = "" if pw == "(en blanco)" else pw
            pwd_menu.add_command(
                label=pw,
                command=lambda p=actual: (
                    self.pass_entry.delete(0, tk.END),
                    self.pass_entry.insert(0, p)
                ))

        tk.Label(self._sql_auth_frame,
                 text="Password del instalador de ContpaqI  \u2014  "
                      "prueba los comunes o pregunta al tecnico que instalo el sistema",
                 bg=C["card"], fg=C["text_mute"], font=("Segoe UI", 8)).grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(0, 4))

        # Indicador de conexion
        self.conn_indicator = tk.Label(
            s1, text="\u25cf  Sin probar",
            bg=C["card"], fg=C["text_mute"],
            font=("Segoe UI", 9, "bold"))
        self.conn_indicator.grid(row=3, column=1, sticky="w", pady=(4, 2))

        self._mk_btn(s1, "\u26a1 Probar conexion", self._test_conn, "ghost").grid(
            row=3, column=2, padx=(6, 0), pady=(4, 2))

        s1.columnconfigure(1, weight=1)

        # ── PASO 2: Empresa ───────────────────────
        s2 = ttk.LabelFrame(body, text="  \U0001f3e2  Paso 2 \u2014 Empresa (base de datos)")
        s2.pack(fill="x", pady=(0, 8))
        tk.Label(s2, text="Empresa:", bg=C["card"], fg=C["text_dim"],
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.db_combo = ttk.Combobox(s2, width=34, state="readonly", font=("Segoe UI", 10))
        self.db_combo.grid(row=0, column=1, sticky="ew", pady=4)
        self.db_combo.bind("<<ComboboxSelected>>", lambda e: self._load_departments())
        self.load_db_btn = self._mk_btn(s2, "\u21bb Cargar empresas", self._load_databases, "primary")
        self.load_db_btn.grid(row=0, column=2, padx=(8, 0))
        s2.columnconfigure(1, weight=1)

        # ── PASO 3: Departamentos ─────────────────
        s3 = ttk.LabelFrame(body,
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

        # ── PASO 4: Filtro toggle ─────────────────
        s4 = ttk.LabelFrame(body, text="  \U0001f3af  Paso 4 \u2014 \u00bfQue empleados exportar?")
        s4.pack(fill="x", pady=(0, 8))
        tr2 = tk.Frame(s4, bg=C["card"])
        tr2.pack(fill="x")
        self._tog_activos = self._mk_toggle(
            tr2, "\u2714  Solo Activos", "Excluye bajas y suspendidos",
            True, lambda: self._set_filter(True))
        self._tog_activos.pack(side="left", fill="both", expand=True, padx=(0, 6), pady=4)
        self._tog_todos = self._mk_toggle(
            tr2, "\u229e  Todos los Empleados", "Incluye bajas, suspendidos e inactivos",
            False, lambda: self._set_filter(False))
        self._tog_todos.pack(side="left", fill="both", expand=True, pady=4)

        # ── PASO 5: Archivo ───────────────────────
        s5 = ttk.LabelFrame(body, text="  \U0001f4be  Paso 5 \u2014 Archivo de salida")
        s5.pack(fill="x", pady=(0, 10))
        tk.Label(s5, text="Ruta:", bg=C["card"], fg=C["text_dim"],
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.file_entry = tk.Entry(
            s5, width=38, font=("Segoe UI", 10),
            bg=C["surface"], fg=C["text"], insertbackground=C["white"],
            relief="flat", highlightthickness=1,
            highlightcolor=C["accent"], highlightbackground=C["border"])
        self.file_entry.grid(row=0, column=1, sticky="ew", pady=4)
        if self.config.get("last_output"):
            self.file_entry.insert(0, self.config["last_output"])
        self._mk_btn(s5, "\U0001f4c1 Examinar", self._select_file, "ghost").grid(
            row=0, column=2, padx=(8, 0))
        s5.columnconfigure(1, weight=1)

        # Progress bar (hidden)
        self.progress = ttk.Progressbar(body, mode="indeterminate")

        # ── Boton GENERAR (verde) ─────────────────
        self.gen_btn = tk.Button(
            body,
            text="  \u25b6   Generar Reporte Excel",
            command=self._execute,
            bg=C["success"], fg=C["white"],
            activebackground="#189B52", activeforeground=C["white"],
            font=("Segoe UI", 12, "bold"),
            relief="flat", cursor="hand2", pady=14)
        self.gen_btn.pack(fill="x", pady=(0, 10))

        # ── Log ───────────────────────────────────
        log_hdr = tk.Frame(body, bg=C["bg"])
        log_hdr.pack(fill="x", pady=(0, 4))
        tk.Label(log_hdr, text="\U0001f4cb  Log de actividad",
                 bg=C["bg"], fg=C["text_dim"],
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        self._mk_btn(log_hdr, "\U0001f4cb Copiar log", self._copy_log, "ghost").pack(side="right")
        self._mk_btn(log_hdr, "\u2715 Limpiar", self._clear_log, "ghost").pack(
            side="right", padx=(0, 6))

        self.log_box = tk.Text(
            body, bg=C["log_bg"], fg="#A8F0C6",
            font=("Consolas", 9), height=12, relief="flat", bd=0,
            state="disabled", wrap="word",
            highlightthickness=1,
            highlightcolor=C["border"], highlightbackground=C["border"])
        self.log_box.pack(fill="x", pady=(0, 6))
        self.log_box.tag_config("info",    foreground="#5BA3FF")
        self.log_box.tag_config("success", foreground=C["success"])
        self.log_box.tag_config("warning", foreground=C["warning"])
        self.log_box.tag_config("error",   foreground="#FF6B6B")
        self.log_box.tag_config("default", foreground="#A8F0C6")
        self.log_box.tag_config("ts",      foreground=C["text_mute"])

        # Status bar
        sb2 = tk.Frame(root, bg=C["surface"], height=28)
        sb2.pack(fill="x", side="bottom")
        sb2.pack_propagate(False)
        self._status_var = tk.StringVar(value="Listo.")
        tk.Label(sb2, textvariable=self._status_var,
                 bg=C["surface"], fg=C["text_dim"],
                 font=("Segoe UI", 9), anchor="w").pack(side="left", padx=10)
        tk.Label(sb2, text="Excel SolutionsV \u00a9 2025",
                 bg=C["surface"], fg=C["text_mute"],
                 font=("Segoe UI", 8)).pack(side="right", padx=10)

    # ── Auth mode ─────────────────────────────────
    def _set_auth_mode(self, sql_auth: bool):
        self._use_sql_auth.set(sql_auth)
        if sql_auth:
            self._auth_sql_btn.configure(bg=C["accent"], fg=C["white"],
                                         font=("Segoe UI", 9, "bold"))
            self._auth_win_btn.configure(bg=C["surface"], fg=C["text_mute"],
                                         font=("Segoe UI", 9))
            self._sql_auth_frame.grid()
        else:
            self._auth_win_btn.configure(bg=C["accent"], fg=C["white"],
                                         font=("Segoe UI", 9, "bold"))
            self._auth_sql_btn.configure(bg=C["surface"], fg=C["text_mute"],
                                         font=("Segoe UI", 9))
            self._sql_auth_frame.grid_remove()
        self._active_driver = None
        self.conn_indicator.configure(text="\u25cf  Sin probar", fg=C["text_mute"])

    def _toggle_pass_visibility(self):
        self.pass_entry.configure(show="" if self._show_pass_var.get() else "\u2022")

    def _get_credentials(self):
        if self._use_sql_auth.get():
            return self.user_entry.get().strip(), self.pass_entry.get()
        return None, None

    # ── Widget helpers ────────────────────────────
    def _mk_btn(self, parent, text, command, kind="primary"):
        cfg = {
            "primary":   (C["accent"],  C["white"],    "#255EAA"),
            "secondary": (C["surface"], C["accent2"],  C["border"]),
            "ghost":     (C["card"],    C["text_dim"], C["border"]),
        }
        bg, fg, abg = cfg[kind]
        return tk.Button(parent, text=text, command=command,
                         bg=bg, fg=fg, activebackground=abg, activeforeground=C["white"],
                         font=("Segoe UI", 9, "bold" if kind == "primary" else "normal"),
                         relief="flat", cursor="hand2", padx=10, pady=5)

    def _mk_toggle(self, parent, label, sublabel, selected, command):
        frame = tk.Frame(parent, bg=C["card"], cursor="hand2",
                         highlightthickness=2,
                         highlightbackground=C["accent"] if selected else C["border"])
        t = tk.Label(frame, text=label, bg=C["card"],
                     fg=C["accent2"] if selected else C["text_dim"],
                     font=("Segoe UI", 11, "bold"), anchor="w", padx=10)
        t.pack(fill="x", pady=(8, 1))
        s = tk.Label(frame, text=sublabel, bg=C["card"],
                     fg=C["text_dim"] if selected else C["text_mute"],
                     font=("Segoe UI", 8), anchor="w", padx=12)
        s.pack(fill="x", pady=(0, 8))
        frame._title_lbl = t
        frame._sub_lbl   = s
        for w in (frame, t, s):
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
    def _set_step(self, step):
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

    def _log(self, msg, level="default"):
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

    # ── Scan ──────────────────────────────────────
    def _start_scan(self):
        if self._scanning:
            return
        self._scanning = True
        self.scan_btn.configure(text="\u23f3 Buscando...", state="disabled")
        self._log("Detectando servidores (red local + unidades mapeadas)...", "info")
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

        # Pre-seleccionar SERVER\COMPAC si existe y no hay nada seleccionado
        if not current:
            for candidate in all_svr:
                if "compac" in candidate.lower() and not candidate.lower().startswith("(") \
                        and not candidate.lower().startswith("local") \
                        and not candidate.startswith("."):
                    self.server_combo.set(candidate)
                    break
            if not self.server_combo.get() and all_svr:
                self.server_combo.set(all_svr[0])

        self.scan_btn.configure(text="\U0001f50d Buscar", state="normal")
        self._log(f"Deteccion completada. {len(servers)} opciones.", "success")
        for s in servers:
            self._log(f"  -> {s}", "default")
        self._log("ContpaqI usa instancia \\COMPAC — busca 'SERVER\\COMPAC' en la lista.", "warning")
        self._set_status(f"Deteccion lista — {len(servers)} opciones.")

    def _on_server_pick(self):
        self._set_step(0)
        self._active_driver = None
        self.conn_indicator.configure(text="\u25cf  Sin probar", fg=C["text_mute"])
        self._log(f"Servidor seleccionado: {self.server_combo.get()}", "info")

    # ── Probar conexion ───────────────────────────
    def _test_conn(self):
        server = self.server_combo.get().strip()
        if not server:
            messagebox.showwarning("Aviso", "Selecciona o escribe el servidor primero.")
            return
        user, password = self._get_credentials()
        if self._use_sql_auth.get() and not user:
            messagebox.showwarning("Aviso", "Escribe el usuario SQL.")
            return

        self.conn_indicator.configure(text="\u25cf  Probando...", fg=C["warning"])
        auth_type = f"SQL auth (usuario: {user})" if user else "Windows auth"
        self._log(f"Probando conexion: {server}  |  {auth_type}", "info")
        self._log(f"Drivers disponibles: {get_installed_sql_drivers()}", "info")
        self._set_status(f"Probando {server}...")

        threading.Thread(
            target=lambda: self.root.after(
                0, self._test_done, server, *test_connection(server, user, password)),
            daemon=True).start()

    def _test_done(self, server, ok, driver, err):
        if ok:
            self._active_driver = driver
            self.conn_indicator.configure(
                text=f"\u25cf  Conectado \u2714  [{driver}]", fg=C["success"])
            self._log(f"Conexion exitosa | Driver: {driver}", "success")
            self._set_status(f"Conectado — {driver}")
            self._set_step(1)
        else:
            self.conn_indicator.configure(text="\u25cf  Fallo", fg=C["danger"])
            self._log(f"Error al conectar: {err}", "error")
            if "does not exist or access denied" in err:
                self._log(
                    "DIAGNOSTICO: 'does not exist or access denied' significa:\n"
                    "  1) El nombre del servidor o instancia es incorrecto, O\n"
                    "  2) El firewall bloquea el puerto 1433 (o UDP 1434), O\n"
                    "  3) El SQL Server Browser no esta activo en el servidor remoto.\n"
                    "  -> Verifica que el servicio se llama 'SQL Server (COMPAC)' y usa SERVER\\COMPAC", "warning")
            elif "Login failed" in err:
                self._log(
                    "DIAGNOSTICO: 'Login failed' = usuario/password incorrecto.\n"
                    "  -> Verifica credenciales con el admin del servidor.", "warning")
            self._set_status("Fallo la conexion. Revisa el log.")

    # ── Cargar empresas ───────────────────────────
    def _load_databases(self):
        server = self.server_combo.get().strip()
        if not server:
            messagebox.showwarning("Aviso", "Selecciona el servidor primero.")
            return
        if not self._active_driver:
            messagebox.showwarning("Aviso", "Primero usa '\u26a1 Probar conexion'.")
            return
        user, password = self._get_credentials()
        self._log(f"Buscando bases NOMINAS en {server}...", "info")
        self._start_progress()
        self.load_db_btn.configure(state="disabled")

        def _work():
            try:
                dbs = get_databases(server, self._active_driver, user, password)
                self.root.after(0, self._dbs_loaded, dbs)
            except Exception as e:
                self.root.after(0, self._on_error, str(e))

        threading.Thread(target=_work, daemon=True).start()

    def _dbs_loaded(self, dbs):
        self._stop_progress()
        self.load_db_btn.configure(state="normal")
        if not dbs:
            self._log("No se encontraron bases NOMINAS*. Verifica la instancia SQL.", "warning")
            messagebox.showwarning("Sin resultados",
                                   "No se encontraron bases de datos NOMINAS* en este servidor.\n\n"
                                   "Verifica que la instancia es la correcta (debe ser \\COMPAC).")
            return
        self.db_combo["values"] = dbs
        self.db_combo.current(0)
        self._log(f"Empresas encontradas: {', '.join(dbs)}", "success")
        self._set_step(2)
        self._load_departments()

    # ── Cargar departamentos ──────────────────────
    def _load_departments(self):
        server = self.server_combo.get().strip()
        db     = self.db_combo.get().strip()
        if not server or not db or not self._active_driver:
            return
        user, password = self._get_credentials()
        self._log(f"Cargando departamentos de '{db}'...", "info")
        self.dept_listbox.delete(0, tk.END)
        self._dept_ids = []

        def _work():
            try:
                df = get_departments(server, db, self._active_driver, user, password)
                self.root.after(0, self._depts_loaded, df)
            except Exception as e:
                self.root.after(0, self._on_error, str(e))

        threading.Thread(target=_work, daemon=True).start()

    def _depts_loaded(self, df):
        for _, row in df.iterrows():
            self._dept_ids.append(row["cIdDepartamento"])
            self.dept_listbox.insert(
                tk.END, f"  {str(row['cIdDepartamento']):>4}  \u2502  {row['cNombreDepartamento']}")
        self._log(f"{len(df)} departamento(s) cargado(s).", "success")
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
        user, password = self._get_credentials()

        missing = []
        if not server:              missing.append("  \u2022 Servidor SQL")
        if not db:                  missing.append("  \u2022 Empresa")
        if not self._active_driver: missing.append("  \u2022 Conexion probada")
        if not file_path:           missing.append("  \u2022 Archivo de salida")
        if missing:
            messagebox.showwarning("Faltan datos", "Completa:\n\n" + "\n".join(missing))
            return

        only_active = self._filter_active
        sel_idx = self.dept_listbox.curselection()
        sel_ids = [self._dept_ids[i] for i in sel_idx] if sel_idx else None

        self._log(
            f"Exportando | Filtro: {'Solo activos' if only_active else 'Todos'} | "
            f"Depts: {f'{len(sel_idx)} seleccionados' if sel_idx else 'todos'}", "info")

        self.gen_btn.configure(state="disabled", text="  \u23f3  Generando...")
        self._start_progress()

        def _work():
            try:
                n = export_data(server, db, file_path, only_active,
                                sel_ids, self._active_driver, user, password)
                self.root.after(0, self._export_done, n, file_path, server, user)
            except Exception as e:
                self.root.after(0, self._on_error, str(e))

        threading.Thread(target=_work, daemon=True).start()

    def _export_done(self, n_rows, file_path, server, user):
        self._stop_progress()
        self.gen_btn.configure(state="normal", text="  \u25b6   Generar Reporte Excel")
        for i, lbl in enumerate(self._step_labels):
            lbl.configure(bg=C["step_done"], fg=C["success"],
                          text=f"  \u2714 {STEPS[i]}  ")
        self._log(f"Exportacion exitosa: {n_rows} empleado(s) -> {file_path}", "success")
        self._set_status(f"\u2714 {n_rows} empleados exportados.")
        save_config({"server": server, "last_output": file_path,
                     "db_user": user or ""})
        messagebox.showinfo("Listo \u2705",
                            f"Exportacion exitosa\n\nEmpleados: {n_rows}\nArchivo: {file_path}")
        if messagebox.askyesno("\u00bfAbrir?", "\u00bfAbrir el Excel ahora?"):
            os.startfile(file_path)

    def _on_error(self, msg):
        self._stop_progress()
        self.gen_btn.configure(state="normal", text="  \u25b6   Generar Reporte Excel")
        self._log(f"ERROR: {msg}", "error")
        self._set_status("\u2718 Error. Revisa el log.")
        messagebox.showerror("Error", f"Ocurrio un error:\n\n{msg}")


# ──────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()