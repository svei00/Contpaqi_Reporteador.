# ==========================
# contpaqi_export_pro.py (ULTRA PRO v3)
# Excel SolutionsV — Exportador de Nóminas ContpaqI
# ==========================

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import subprocess
import pyodbc
import pandas as pd
import json
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

CONFIG_FILE = "config.json"

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
    "success":   "#2ECC71",
    "danger":    "#E74C3C",
    "warning":   "#F39C12",
    "text":      "#E8ECF0",
    "text_dim":  "#7B8499",
    "text_mute": "#454B63",
    "white":     "#FFFFFF",
}

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
# AUTO-DETECCIÓN DE SERVIDORES SQL SERVER
# ──────────────────────────────────────────────
def detect_sql_servers():
    """
    Detecta instancias de SQL Server en la red usando tres métodos:
    1. sqlcmd -L  (lista servidores visibles en la red local)
    2. pyodbc con SQL Server Browser (UDP 1434)
    3. Registro de Windows (instancias locales)
    Retorna una lista de strings con los nombres/IPs encontrados.
    """
    found = set()

    # Método 1: sqlcmd -L
    try:
        result = subprocess.run(
            ["sqlcmd", "-L"],
            capture_output=True, text=True, timeout=8
        )
        for line in result.stdout.splitlines():
            s = line.strip()
            if s and not s.lower().startswith("server"):
                found.add(s)
    except Exception:
        pass

    # Método 2: osql -L (alternativo, disponible en SQL Server 2008+)
    try:
        result = subprocess.run(
            ["osql", "-L"],
            capture_output=True, text=True, timeout=8
        )
        for line in result.stdout.splitlines():
            s = line.strip()
            if s and not s.lower().startswith("server"):
                found.add(s)
    except Exception:
        pass

    # Método 3: Registro de Windows — instancias locales
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

    # Siempre incluir opciones comunes como fallback
    defaults = ["(local)", "localhost", r"localhost\SQLEXPRESS", r".\SQLEXPRESS"]
    for d in defaults:
        found.add(d)

    return sorted(found)


# ──────────────────────────────────────────────
# BASE DE DATOS
# ──────────────────────────────────────────────
def test_connection(server):
    """Prueba si el servidor responde. Retorna True/False."""
    try:
        conn_str = (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={server};Trusted_Connection=yes;"
            "Connect Timeout=5;"
        )
        conn = pyodbc.connect(conn_str)
        conn.close()
        return True
    except Exception:
        return False


def get_databases(server):
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={server};Trusted_Connection=yes;"
        "Connect Timeout=10;"
    )
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sys.databases WHERE name LIKE 'NOMINAS%' ORDER BY name")
    rows = [row[0] for row in cursor.fetchall()]
    conn.close()
    return rows


def get_departments(server, database):
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={server};DATABASE={database};Trusted_Connection=yes;"
        "Connect Timeout=10;"
    )
    conn = pyodbc.connect(conn_str)
    query = "SELECT cIdDepartamento, cNombreDepartamento FROM NomDepartamento ORDER BY cNombreDepartamento"
    df = pd.read_sql(query, conn)
    conn.close()
    return df


# ──────────────────────────────────────────────
# QUERY
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
        E.cCodigoEmpleado   AS Codigo,
        E.cRFC              AS RFC,
        E.cCURP             AS CURP,
        E.cNombre           AS Nombre,
        E.cSueldoDiario     AS [Salario Diario],
        E.cSalarioDiarioIntegrado AS SDI,
        D.cNombreDepartamento     AS Departamento
    FROM NomEmpleado E
    LEFT JOIN NomDepartamento D
        ON E.cIdDepartamento = D.cIdDepartamento
    {where_clause}
    ORDER BY E.cCodigoEmpleado
    """


# ──────────────────────────────────────────────
# EXPORTAR
# ──────────────────────────────────────────────
def export_data(server, database, output_path, only_active, selected_departments):
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={server};DATABASE={database};Trusted_Connection=yes;"
        "Connect Timeout=15;"
    )
    conn = pyodbc.connect(conn_str)
    query = build_query(only_active, selected_departments)
    df = pd.read_sql(query, conn)
    conn.close()

    if df.empty:
        raise ValueError("La consulta no arrojó resultados. Verifica los filtros seleccionados.")

    df.to_excel(output_path, index=False, engine="openpyxl")

    wb = load_workbook(output_path)
    ws = wb.active
    ws.title = "Empleados"

    BLUE     = "3182DF"
    DARK     = "0F1117"
    ROW_ALT  = "EEF4FB"
    ROW_EVEN = "FFFFFF"
    BORDER_C = "BDD0E8"

    thin_border = Border(
        left=Side(style="thin",   color=BORDER_C),
        right=Side(style="thin",  color=BORDER_C),
        top=Side(style="thin",    color=BORDER_C),
        bottom=Side(style="thin", color=BORDER_C),
    )

    # — Encabezados
    for cell in ws[1]:
        cell.fill      = PatternFill(start_color=BLUE, end_color=BLUE, fill_type="solid")
        cell.font      = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
        cell.border    = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[1].height = 24

    # — Filas de datos
    max_row = ws.max_row
    for row_idx in range(2, max_row + 1):
        fill_color = ROW_ALT if row_idx % 2 == 0 else ROW_EVEN
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill   = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
            cell.border = thin_border
            cell.font   = Font(name="Calibri", size=10)
            cell.alignment = Alignment(vertical="center")

    # — Ancho de columnas automático
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 4, 50)

    # — Fila de totales
    total_row = max_row + 2
    tc = ws.cell(row=total_row, column=1, value="TOTAL EMPLEADOS:")
    tc.font      = Font(bold=True, color=DARK, name="Calibri", size=11)
    tc.fill      = PatternFill(start_color="D6E4F5", end_color="D6E4F5", fill_type="solid")
    tc.alignment = Alignment(horizontal="right")

    vc = ws.cell(row=total_row, column=2, value=max_row - 1)
    vc.font      = Font(bold=True, color=BLUE, name="Calibri", size=11)
    vc.fill      = PatternFill(start_color="D6E4F5", end_color="D6E4F5", fill_type="solid")
    vc.alignment = Alignment(horizontal="center")

    # — Freeze panes (encabezado siempre visible)
    ws.freeze_panes = "A2"

    wb.save(output_path)
    return len(df)


# ──────────────────────────────────────────────
# GUI
# ──────────────────────────────────────────────
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel SolutionsV · Exportador Nóminas ContpaqI")
        self.root.configure(bg=C["bg"])
        self.root.resizable(False, False)

        self.config = load_config()
        self._dept_ids = []          # IDs paralelos al listbox
        self._scanning = False

        self._apply_styles()
        self._build_ui()
        self._center_window(740, 640)

        # Auto-scan al arrancar
        self.root.after(300, self._start_scan)

    # ── Estilos ttk ──────────────────────────────
    def _apply_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        # Frame / LabelFrame
        style.configure("TFrame",      background=C["bg"])
        style.configure("Card.TFrame", background=C["card"])

        style.configure(
            "TLabelframe",
            background=C["card"], foreground=C["text_dim"],
            bordercolor=C["border"], relief="flat", padding=12,
        )
        style.configure(
            "TLabelframe.Label",
            background=C["card"], foreground=C["accent2"],
            font=("Segoe UI", 9, "bold"),
        )

        # Combobox
        style.configure(
            "TCombobox",
            fieldbackground=C["surface"], background=C["surface"],
            foreground=C["text"], selectbackground=C["accent"],
            selectforeground=C["white"], bordercolor=C["border"],
            arrowcolor=C["accent2"],
        )
        style.map("TCombobox",
            fieldbackground=[("readonly", C["surface"])],
            foreground=[("readonly", C["text"])],
        )

        # Entry
        style.configure(
            "TEntry",
            fieldbackground=C["surface"], foreground=C["text"],
            bordercolor=C["border"], insertcolor=C["text"],
        )

        # Radiobutton
        style.configure(
            "TRadiobutton",
            background=C["card"], foreground=C["text"],
            font=("Segoe UI", 10),
        )
        style.map("TRadiobutton",
            background=[("active", C["card"])],
            foreground=[("active", C["accent2"])],
        )

        # Scrollbar
        style.configure(
            "Vertical.TScrollbar",
            background=C["surface"], troughcolor=C["bg"],
            arrowcolor=C["text_dim"], bordercolor=C["border"],
        )

        # Separator
        style.configure("TSeparator", background=C["border"])

    # ── Construcción de UI ────────────────────────
    def _build_ui(self):
        root = self.root

        # ─ Título superior ─
        header = tk.Frame(root, bg=C["accent"], height=54)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(
            header,
            text="  ⬡  Excel SolutionsV",
            bg=C["accent"], fg=C["white"],
            font=("Segoe UI", 15, "bold"),
            anchor="w",
        ).pack(side="left", padx=20, pady=10)

        tk.Label(
            header,
            text="Exportador de Nóminas ContpaqI",
            bg=C["accent"], fg="#BFD9FF",
            font=("Segoe UI", 10),
        ).pack(side="left", padx=4, pady=10)

        # ─ Cuerpo ─
        body = tk.Frame(root, bg=C["bg"], padx=18, pady=16)
        body.pack(fill="both", expand=True)

        # ── Tarjeta: Conexión ──
        conn_frame = ttk.LabelFrame(body, text="  🔌  Conexión al Servidor SQL")
        conn_frame.pack(fill="x", pady=(0, 12))
        conn_frame.configure(style="TLabelframe")

        tk.Label(conn_frame, text="Servidor:", bg=C["card"], fg=C["text_dim"],
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", padx=(0,8), pady=4)

        self.server_combo = ttk.Combobox(conn_frame, width=38, state="normal", font=("Segoe UI", 10))
        self.server_combo.grid(row=0, column=1, sticky="ew", pady=4)
        saved_server = self.config.get("server", "")
        if saved_server:
            self.server_combo.set(saved_server)

        self.scan_btn = tk.Button(
            conn_frame, text="🔍 Buscar",
            command=self._start_scan,
            bg=C["surface"], fg=C["accent2"],
            activebackground=C["border"], activeforeground=C["white"],
            font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2",
            padx=10, pady=4,
        )
        self.scan_btn.grid(row=0, column=2, padx=(8,0), pady=4)

        self.ping_label = tk.Label(conn_frame, text="", bg=C["card"],
                                   font=("Segoe UI", 9))
        self.ping_label.grid(row=0, column=3, padx=(8,0))

        # Botón probar conexión
        tk.Button(
            conn_frame, text="⚡ Probar conexión",
            command=self._test_conn,
            bg=C["surface"], fg=C["text_dim"],
            activebackground=C["border"], activeforeground=C["white"],
            font=("Segoe UI", 9), relief="flat", cursor="hand2",
            padx=8, pady=4,
        ).grid(row=1, column=1, sticky="w", pady=(0,4))

        conn_frame.columnconfigure(1, weight=1)

        # ── Tarjeta: Empresa ──
        emp_frame = ttk.LabelFrame(body, text="  🏢  Empresa (Base de Datos)")
        emp_frame.pack(fill="x", pady=(0, 12))

        tk.Label(emp_frame, text="Empresa:", bg=C["card"], fg=C["text_dim"],
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", padx=(0,8), pady=4)

        self.db_combo = ttk.Combobox(emp_frame, width=38, state="readonly", font=("Segoe UI", 10))
        self.db_combo.grid(row=0, column=1, sticky="ew", pady=4)
        self.db_combo.bind("<<ComboboxSelected>>", lambda e: self._load_departments())

        self.load_db_btn = tk.Button(
            emp_frame, text="↻ Cargar",
            command=self._load_databases,
            bg=C["accent"], fg=C["white"],
            activebackground=C["accent2"], activeforeground=C["white"],
            font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2",
            padx=10, pady=4,
        )
        self.load_db_btn.grid(row=0, column=2, padx=(8,0), pady=4)
        emp_frame.columnconfigure(1, weight=1)

        # ── Tarjeta: Departamentos ──
        dept_frame = ttk.LabelFrame(body, text="  📂  Departamentos  (Ctrl+clic para selección múltiple)")
        dept_frame.pack(fill="x", pady=(0, 12))

        list_container = tk.Frame(dept_frame, bg=C["card"])
        list_container.pack(fill="x")

        self.dept_listbox = tk.Listbox(
            list_container,
            selectmode=tk.MULTIPLE,
            bg=C["surface"], fg=C["text"],
            selectbackground=C["accent"], selectforeground=C["white"],
            font=("Consolas", 10),
            height=6, relief="flat", bd=0,
            activestyle="none",
            highlightthickness=1, highlightcolor=C["border"],
            highlightbackground=C["border"],
        )
        self.dept_listbox.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(list_container, orient="vertical",
                                  command=self.dept_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.dept_listbox.configure(yscrollcommand=scrollbar.set)

        dept_hint = tk.Label(dept_frame,
                             text="Sin selección = todos los departamentos",
                             bg=C["card"], fg=C["text_mute"],
                             font=("Segoe UI", 8))
        dept_hint.pack(anchor="w", pady=(4,0))

        # ── Tarjeta: Archivo de salida ──
        file_frame = ttk.LabelFrame(body, text="  💾  Archivo de Salida")
        file_frame.pack(fill="x", pady=(0, 12))

        tk.Label(file_frame, text="Ruta:", bg=C["card"], fg=C["text_dim"],
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", padx=(0,8), pady=4)

        self.file_entry = ttk.Entry(file_frame, width=40, font=("Segoe UI", 10))
        self.file_entry.grid(row=0, column=1, sticky="ew", pady=4)
        saved_path = self.config.get("last_output", "")
        if saved_path:
            self.file_entry.insert(0, saved_path)

        tk.Button(
            file_frame, text="📁 Examinar",
            command=self._select_file,
            bg=C["surface"], fg=C["text"],
            activebackground=C["border"], activeforeground=C["white"],
            font=("Segoe UI", 9), relief="flat", cursor="hand2",
            padx=8, pady=4,
        ).grid(row=0, column=2, padx=(8,0), pady=4)
        file_frame.columnconfigure(1, weight=1)

        # ── Filtro de estatus ──
        filter_frame = ttk.LabelFrame(body, text="  🎯  Filtro de Empleados")
        filter_frame.pack(fill="x", pady=(0, 16))

        self.status_var = tk.StringVar(value="activos")
        opts = [("Solo empleados activos", "activos"), ("Todos los empleados", "todos")]
        for i, (label, val) in enumerate(opts):
            ttk.Radiobutton(
                filter_frame, text=label,
                variable=self.status_var, value=val,
            ).grid(row=0, column=i, padx=16, pady=6, sticky="w")

        # ── Botón principal + barra de progreso ──
        self.gen_btn = tk.Button(
            body,
            text="  ▶  Generar Reporte Excel",
            command=self._execute,
            bg=C["accent"], fg=C["white"],
            activebackground="#255EAA", activeforeground=C["white"],
            font=("Segoe UI", 12, "bold"),
            relief="flat", cursor="hand2",
            pady=12,
        )
        self.gen_btn.pack(fill="x", pady=(0, 8))

        self.progress = ttk.Progressbar(body, mode="indeterminate",
                                        style="TProgressbar")
        # (oculto hasta que se use)

        # ── Barra de estado ──
        status_bar = tk.Frame(root, bg=C["surface"], height=28)
        status_bar.pack(fill="x", side="bottom")
        status_bar.pack_propagate(False)

        self.status_var_lbl = tk.StringVar(value="Listo.")
        tk.Label(
            status_bar,
            textvariable=self.status_var_lbl,
            bg=C["surface"], fg=C["text_dim"],
            font=("Segoe UI", 9), anchor="w",
        ).pack(side="left", padx=12, pady=4)

        tk.Label(
            status_bar,
            text="Excel SolutionsV © 2025",
            bg=C["surface"], fg=C["text_mute"],
            font=("Segoe UI", 8),
        ).pack(side="right", padx=12, pady=4)

    # ── Helpers UI ────────────────────────────────
    def _center_window(self, w, h):
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x  = (sw - w) // 2
        y  = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _set_status(self, msg, color=None):
        self.status_var_lbl.set(msg)
        if color:
            for w in self.root.winfo_children():
                pass  # solo actualizamos el label de status (ya está en StringVar)

    def _start_progress(self):
        self.progress.pack(fill="x", pady=(0, 8))
        self.progress.start(10)

    def _stop_progress(self):
        self.progress.stop()
        self.progress.pack_forget()

    # ── Scan de servidores ────────────────────────
    def _start_scan(self):
        if self._scanning:
            return
        self._scanning = True
        self.scan_btn.configure(text="⏳ Buscando…", state="disabled")
        self._set_status("Detectando servidores SQL Server en la red…")
        threading.Thread(target=self._do_scan, daemon=True).start()

    def _do_scan(self):
        servers = detect_sql_servers()
        self.root.after(0, self._scan_done, servers)

    def _scan_done(self, servers):
        self._scanning = False
        current = self.server_combo.get()
        self.server_combo["values"] = servers
        if current and current not in servers:
            servers_list = list(servers) + [current]
            self.server_combo["values"] = servers_list
        if not current and servers:
            self.server_combo.set(servers[0])
        self.scan_btn.configure(text="🔍 Buscar", state="normal")
        self._set_status(f"Se encontraron {len(servers)} servidor(es). Selecciona uno y carga las empresas.")

    # ── Probar conexión ───────────────────────────
    def _test_conn(self):
        server = self.server_combo.get().strip()
        if not server:
            messagebox.showwarning("Aviso", "Selecciona o escribe el nombre del servidor primero.")
            return
        self._set_status(f"Probando conexión a {server}…")
        self.ping_label.configure(text="…", fg=C["text_dim"])

        def do_test():
            ok = test_connection(server)
            self.root.after(0, self._test_done, ok, server)

        threading.Thread(target=do_test, daemon=True).start()

    def _test_done(self, ok, server):
        if ok:
            self.ping_label.configure(text="✔ Conectado", fg=C["success"])
            self._set_status(f"Conexión exitosa a {server}.")
        else:
            self.ping_label.configure(text="✘ Sin respuesta", fg=C["danger"])
            self._set_status(f"No se pudo conectar a {server}. Verifica el servidor y permisos.")

    # ── Cargar bases de datos ─────────────────────
    def _load_databases(self):
        server = self.server_combo.get().strip()
        if not server:
            messagebox.showwarning("Aviso", "Escribe o selecciona el servidor primero.")
            return

        self._set_status("Cargando empresas…")
        self._start_progress()
        self.load_db_btn.configure(state="disabled")

        def do_load():
            try:
                dbs = get_databases(server)
                self.root.after(0, self._dbs_loaded, dbs)
            except Exception as e:
                self.root.after(0, self._show_error, str(e))

        threading.Thread(target=do_load, daemon=True).start()

    def _dbs_loaded(self, dbs):
        self._stop_progress()
        self.load_db_btn.configure(state="normal")
        if not dbs:
            messagebox.showwarning("Sin resultados",
                                   "No se encontraron bases de datos NOMINAS* en este servidor.")
            self._set_status("Sin bases de datos NOMINAS encontradas.")
            return
        self.db_combo["values"] = dbs
        self.db_combo.current(0)
        self._set_status(f"{len(dbs)} empresa(s) encontrada(s). Selecciona una.")
        self._load_departments()

    # ── Cargar departamentos ──────────────────────
    def _load_departments(self):
        server = self.server_combo.get().strip()
        db     = self.db_combo.get().strip()
        if not server or not db:
            return

        self._set_status("Cargando departamentos…")
        self.dept_listbox.delete(0, tk.END)
        self._dept_ids = []

        def do_load():
            try:
                df = get_departments(server, db)
                self.root.after(0, self._depts_loaded, df)
            except Exception as e:
                self.root.after(0, self._show_error, str(e))

        threading.Thread(target=do_load, daemon=True).start()

    def _depts_loaded(self, df):
        for _, row in df.iterrows():
            self._dept_ids.append(row["cIdDepartamento"])
            self.dept_listbox.insert(
                tk.END, f"  {row['cIdDepartamento']:>4}  │  {row['cNombreDepartamento']}"
            )
        self._set_status(f"{len(df)} departamento(s) cargado(s). Selecciona o deja vacío para todos.")

    # ── Seleccionar archivo ───────────────────────
    def _select_file(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
            title="Guardar reporte como…",
        )
        if path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, path)

    # ── Ejecutar exportación ──────────────────────
    def _execute(self):
        server    = self.server_combo.get().strip()
        db        = self.db_combo.get().strip()
        file_path = self.file_entry.get().strip()

        if not server:
            messagebox.showwarning("Falta dato", "Selecciona el servidor SQL.")
            return
        if not db:
            messagebox.showwarning("Falta dato", "Selecciona la empresa.")
            return
        if not file_path:
            messagebox.showwarning("Falta dato", "Elige dónde guardar el archivo Excel.")
            return

        only_active = self.status_var.get() == "activos"
        sel_idx     = self.dept_listbox.curselection()
        sel_ids     = [self._dept_ids[i] for i in sel_idx] if sel_idx else None

        self.gen_btn.configure(state="disabled", text="  ⏳  Generando…")
        self._start_progress()
        self._set_status("Exportando datos, por favor espera…")

        def do_export():
            try:
                n = export_data(server, db, file_path, only_active, sel_ids)
                self.root.after(0, self._export_done, n, file_path, server, file_path)
            except Exception as e:
                self.root.after(0, self._show_error, str(e))

        threading.Thread(target=do_export, daemon=True).start()

    def _export_done(self, n_rows, file_path, server, last_output):
        self._stop_progress()
        self.gen_btn.configure(state="normal", text="  ▶  Generar Reporte Excel")
        self._set_status(f"✔ Reporte generado: {n_rows} empleado(s)  →  {file_path}")

        save_config({"server": server, "last_output": last_output})

        messagebox.showinfo(
            "Reporte generado",
            f"✅ Exportación exitosa\n\n"
            f"  Empleados:  {n_rows}\n"
            f"  Archivo:    {file_path}\n\n"
            "Puedes abrir el archivo ahora.",
        )

        # Ofrecer abrir el archivo
        if messagebox.askyesno("Abrir archivo", "¿Deseas abrir el archivo Excel ahora?"):
            os.startfile(file_path)

    def _show_error(self, msg):
        self._stop_progress()
        self.gen_btn.configure(state="normal", text="  ▶  Generar Reporte Excel")
        self._set_status(f"✘ Error: {msg}")
        messagebox.showerror("Error", f"Ocurrió un error:\n\n{msg}")


# ──────────────────────────────────────────────
# ENTRADA
# ──────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    app  = App(root)
    root.mainloop()