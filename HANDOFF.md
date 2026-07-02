# HANDOFF — contpaqi_reporteador

> **Documento maestro de arquitectura y roadmap.** Escrito por la sesión de
> auditoría (Fable). La implementación la hará Sonnet 5 u Opus 4.8 siguiendo
> este archivo. Es **standalone**: no depende de ningún transcript previo.
>
> **Idioma del proyecto:** español (usuario y usuarios finales mexicanos).
> **Fecha de auditoría:** 2026-07-10.
> **Repo auditado:** `D:\repos\contpaqi_reporteador`
> **Estado del código auditado:** rama `main`, commit `58aca0f`.
>
> **Revisión 2 (2026-07-10):** refinamiento de 4 puntos tras feedback de Svei.
> (a) Memoria de última carpeta de guardado por perfil + corrección del bug de
> "memoria fantasma" de servidor/usuario. (b) Profundidad operativa del
> almacenamiento de contraseña (fallback de `keyring` sin backend seguro + flujo
> de UX) **y reencuadre del "Auto-Buscar": NO se elimina, se conserva como
> recuperación de la contraseña *default* de ContpaqI, corrigiendo solo su
> almacenamiento y recortando entradas genéricas.** (c) Regla dura de código
> Python autoexplicativo + comentarios persistentes, y tabla de mapeo
> monolito→módulos. (d) `.gitignore` creado como archivo real. Los cambios de
> esta revisión están marcados con «**[R2]**» en cada sección afectada.

---

## 0. Cómo leer este documento

1. **Sección 1** — Hallazgos de la auditoría (Fase 0), con evidencia
   (`archivo:línea`). Léela primero: justifica todo lo demás.
2. **Sección 2** — Lo que Svei confirmó sobre alcance y visión. Esto
   reencuadró el proyecto: de "auditar un exportador de un reporte" a
   "diseñar un motor de reportes de nómina escalable".
3. **Sección 3** — Las 5 decisiones de arquitectura (Fases 1–5), ya
   resueltas, con justificación.
4. **Sección 4** — Catálogo de reportes propuesto (~35 reportes).
5. **Sección 5** — Arquitectura de módulos objetivo.
6. **Sección 6** — Generador de recibos personalizable + subcontratados.
7. **Sección 7** — Fórmula de suavizado de color (tint), explícita.
8. **Sección 8** — Sistema de temas (8 variantes con tipografía).
9. **Sección 9** — Roadmap por fases con criterios de salida.
10. **Sección 10** — Acciones inmediatas de seguridad/git (hacer ANTES de codear).
11. **Sección 11** — Reglas de estilo de código (íntegras, de Svei).
12. **Sección 12** — Git workflow.

> **Nota sobre el "roadmap" pedido en Fase 0:** el README ya existe (está en
> UTF-16, por eso puede verse con espacios en algunos editores). El roadmap
> NO existía como archivo; vive ahora en la **Sección 9** de este documento.
> No se creó un `ROADMAP.md` aparte para no fragmentar la fuente de verdad.

---

## 1. Fase 0 — Hallazgos de la auditoría (con evidencia)

### 1.1 Arquitectura actual: un solo archivo monolítico

Todo el proyecto es **un único script de 1,034 líneas**:
`contpaqi_exporter.py`. Ese archivo mezcla, sin separación:

- Capa de datos (conexión ODBC, introspección de esquema, consultas SQL).
- Lógica de negocio (armado de RFC/CURP, filtros, construcción del Excel).
- Presentación (toda la GUI Tkinter, ~600 líneas dentro de la clase `App`).
- Estilos de salida hardcodeados dentro de la función de exportación.

Archivos del repo:
```
contpaqi_exporter.py   # 1034 líneas — TODO el programa
config.json            # { "known_passwords": ["Compac1"] }  ← ver 1.4
README.md              # existe, en UTF-16
run.bat / run.sh       # lanzadores
Trabajadores Activos al *.xlsx   # 10 reportes generados, sueltos ← ver 1.5
```

No hay: `requirements.txt` / `pyproject.toml`, `.gitignore`, carpeta de
tests, ni módulos separados. Las dependencias solo se mencionan en prosa
del README (`pyodbc`, `pandas`, `openpyxl`).

**Implicación:** esta estructura NO escala a un catálogo de ~35 reportes con
filtros dinámicos, temas y plantillas. La modularización (Sección 5) es
prerrequisito de todo lo demás.

### 1.2 Framework GUI: Tkinter puro (no ttkbootstrap, no PySide6)

`import tkinter as tk` + `ttk` con tema `"clam"`
(`contpaqi_exporter.py:6-7`, `:469`). El aspecto oscuro es un **tema hecho a
mano**: un diccionario de ~18 colores hex (`C = {...}`,
`contpaqi_exporter.py:47-65`) aplicado widget por widget. No usa librería de
temas. Esto significa: sin light/dark nativo, cada widget re-estilizado a
mano, y ~600 líneas de UI imperativa difícil de extender.

### 1.3 Conexión a datos: pyodbc directo a SQL Server (confirmado)

- `import pyodbc` (`contpaqi_exporter.py:12`).
- Autodetección de driver por lista de prioridad
  (`ODBC_DRIVERS_PRIORITY`, `:31-38`) y `pyodbc.drivers()` (`:40-42`).
- Cadena de conexión construida en `_build_cs` (`:179-191`): soporta
  `Trusted_Connection=yes` (Windows) o `UID/PWD` (SQL).
- Introspección dinámica de esquema para tolerar versiones viejas vs.
  nuevas de ContpaqI: descubre nombres reales de tablas y columnas vía
  `sys.tables` e `INFORMATION_SCHEMA.COLUMNS`
  (`get_departments:223-251`, `export_data:260-373`). El helper
  `get_real_col` (`:280-285`) mapea nombres lógicos → físicos probando
  candidatos. **Este patrón es bueno y hay que conservarlo y formalizarlo**
  (ver Decisión 3 y Sección 5).
- Reporte único hoy: empleados. Tablas `nom10001`/`NomEmpleado` (empleados)
  y `nom10003`/`NomDepartamento` (departamentos). Columnas de salida:
  Código, RFC, CURP, Nombre, Salario Diario, SDI, Departamento
  (query en `:354-367`).

### 1.4 Manejo de contraseña — HALLAZGO CRÍTICO

Tres problemas encadenados:

1. **Texto plano en disco.** `config.json` guarda
   `"known_passwords": ["Compac1"]` (`config.json:1-5`). `save_known_password`
   escribe la clave exitosa sin cifrar (`contpaqi_exporter.py:101-110`).

2. **Diccionario de contraseñas hardcodeado + botón "Auto-Buscar".** Hay una
   lista de 20 contraseñas (`DEFAULT_PASSWORDS`, `:84-89`) y un botón
   "⚡ Auto-Buscar" que las itera contra el SQL Server hasta acertar
   (`_auto_check_passwords:860-904`).
   **[R2] — Corrección de la valoración de la revisión 1.** La sesión de
   auditoría original marcó esto como "ataque de diccionario, eliminar por
   completo". Svei corrigió el encuadre y tiene razón: el caso de uso es
   **legítimo**. El contador administra el servidor SQL de ContpaqI de su
   cliente, pero **quien instala ContpaqI a menudo NO le entrega la contraseña
   default del SQL** (para cobrarle horas de soporte cuando la necesite). Esas
   contraseñas default de ContpaqI (`Compac1`, etc.) son **públicas y
   documentadas**, y el usuario está autorizado en ese servidor. Recuperar el
   acceso a una base que uno administra usando la contraseña default
   documentada del propio producto **no es fuerza bruta maliciosa** — es una
   comodidad válida y un diferenciador real para contadores.
   **Qué sí hay que corregir (no eliminar):** (i) el resultado ya no se guarda
   en texto plano sino en `keyring` (ver Decisión 1); (ii) recortar la lista a
   las contraseñas default **reales de ContpaqI**, quitando las genéricas que
   no son de ContpaqI (`123456`, `admin`, `Admin123`, `root`, `1234`, `sql`,
   `server`, `master`, `temporal`…) — esas no aportan al caso de uso y son
   justo lo que hace que parezca herramienta de ataque; (iii) dejar claro en la
   UI que solo corre contra el servidor que el propio usuario seleccionó.

3. **La contraseña está commiteada a git.** `config.json` está trackeado
   (`git ls-files config.json` → sí) y `Compac1` entró al historial en el
   commit `4620e7c`. Es la contraseña *default* pública de ContpaqI (baja
   sensibilidad real), pero el patrón es inaceptable y hay que corregir
   antes de crecer. Ver acciones en Sección 10.

### 1.5 PII expuesta sin .gitignore

Los 10 archivos `Trabajadores Activos al *.xlsx` (con RFC, CURP y salarios de
empleados reales) están sueltos en la carpeta del repo. `git status` los
muestra como *untracked*, pero **no existe `.gitignore`**, así que un
`git add .` los subiría. Riesgo de fuga de datos personales. Ver Sección 10.

### 1.6 README y roadmap

- `README.md` **existe** (UTF-16 LE; por eso algunos editores lo muestran
  con espacios entre caracteres). Cubre requisitos, instalación y uso paso a
  paso. Un desajuste menor: el FAQ dice que la app "usa autenticación de
  Windows (Trusted Connection)" pero el modo por defecto en código es SQL
  (`self._auth_mode = "sql"`, `:451`). Corregir el texto o el default.
- **Roadmap: no existía.** Se entrega en la Sección 9.

### 1.7 Otros hallazgos de calidad (código construido con Gemini free)

Revisados con el mismo criterio que cualquier auditoría; ni asumiendo que
está mal ni que está bien:

| # | Hallazgo | Evidencia | Severidad |
|---|----------|-----------|-----------|
| a | Modo SDK es decorativo/muerto: `sdk_test_connection` siempre retorna `False`; `_test_done_sdk` siempre reporta fallo. Feature que no hace nada. | `:132-141`, `:949-952` | Media |
| b | Manejo de errores inconsistente: solo `_execute` envuelve en try/except (`:1002-1007`). `_load_databases`/`_load_departments` corren en thread SIN captura; si la conexión cae a mitad, la excepción muere silenciosa en el hilo. | `:954-958`, `:970-979` | Media |
| c | `except:` desnudos que se tragan todo (incluye KeyboardInterrupt). | `:154`, `:247` | Baja |
| d | IDs de departamento interpolados directo en SQL vía f-string (`",".join(map(str, ...))`). Vienen de la BD (riesgo bajo), pero el patrón es inseguro y debe parametrizarse. Nombres de tabla/columna también se interpolan (necesario por el esquema dinámico, pero deben validarse contra el whitelist introspectado). | `:347-349`, `:237` | Media |
| e | Estilos de Excel hardcodeados dentro de `export_data` (`BLUE`, `ROW_ALT`, bordes). Debe volverse el sistema de temas (Sección 8). | `:384-434` | (rediseño) |
| f | Heurística frágil de "activo": decide `IN ('A','R')` vs `= 1` según si el nombre de la columna contiene "estado". Debe mapearse explícito por versión. | `:345-346` | Media |
| g | `safe_str` definida pero nunca usada (código muerto). | `:256-258` | Baja |
| h | Sin `requirements.txt`/`pyproject.toml`; deps solo en prosa del README. Sin pins. | — | Media |
| i | Escritura de `config.json` no atómica; dos perfiles concurrentes podrían pisarse. Relevante al pasar a multi-perfil. | `:78-81` | Baja |

**Lo que SÍ está bien y hay que conservar:** la introspección dinámica de
esquema (`get_real_col`), el uso de threads para no congelar la UI, el panel
de log con timestamps, el autofilter + fila de totales con
`SUBTOTAL(103,...)` (cuenta solo filas visibles al filtrar), y el naming
dinámico del archivo de salida por fecha y filtro.

---

## 2. Alcance y visión confirmados por Svei

Svei reencuadró el proyecto en la sesión de auditoría. Puntos que mandan
sobre el diseño:

1. **Multi-empresa AISLADA.** Necesita varias empresas/RFC en paralelo, pero
   **"no se deben ver ni mezclar"**. Ya trabaja con dos empresas en la misma
   app sin problema. → Perfiles estrictamente separados.

2. **No quiere "minucias", quiere MUCHOS reportes nuevos.** Las variantes
   "con/sin SDI, días laborados, departamento" **son vistas del mismo
   reporte**, no reportes nuevos. Pidió explícitamente proponer **muchos**
   reportes de nómina aprovechables desde ContpaqI Nóminas, con foco en
   **escalabilidad**. Ideas suyas textuales:
   - Un reporte que traiga **todas las cantidades** + una hoja con
     **plantilla para imprimir los recibos a su gusto**, personalizados con
     logo y colores institucionales (el "sobre recibo" de ContpaqI es lento).
   - Soportar **trabajadores fuera de ContpaqI** (subcontratados): capturar
     sus datos ahí mismo y tener un reporte 100% personalizable con los
     mismos factores.

3. **Contabilidad/Balanza/EdoR: a futuro.** Coquetea con conectar el sistema
   de nómina con el de contabilidad/facturación, pero acepta mantenerlos
   separados. Le atrae el valor agregado de **hacer reportes personalizados y
   predefinidos desde esta interfaz sin abrir el reporteador de ContpaqI**.

4. **Potencial de venta: incierto pero optimista.** Si se vuelve un "monstruo
   reporteador" y es personalizable, y como muchos despachos usan ContpaqI,
   ve valor comercial. → Diseñar branding/marca blanca de forma que **no
   estorbe** y quede listo para vender, sin sobre-invertir hoy.

---

## 3. Las 5 decisiones de arquitectura (Fases 1–5), resueltas

### Decisión 1 (Fase 1) — Perfiles multi-empresa aislados + credenciales en Windows Credential Manager

**Qué:** reemplazar `config.json` por un modelo de **perfiles**, uno por
empresa/RFC, con la contraseña **fuera del JSON**.

- `profiles.json` (no versionado) guarda solo metadatos **no secretos** por
  perfil: `id`, `nombre_visible`, `servidor`, `base_de_datos`, `driver_odbc`,
  `usuario_sql`, `modo_auth` (`windows`/`sql`), los últimos filtros usados y
  **`ultima_carpeta_guardado`** (ver 1.a más abajo).
- **La contraseña se guarda en el Windows Credential Manager** vía la
  librería `keyring`, con clave `contpaqi_reporteador/<profile_id>`.

**Por qué keyring/Credential Manager y no Fernet:** Fernet exige guardar y
custodiar una llave; eso solo mueve el problema ("¿dónde guardo la llave que
cifra la contraseña?"). El Credential Manager de Windows ya cifra por usuario
del SO, sin archivo de llave ni master password, y la app es Windows-only. Es
la opción que mejor encaja con el stack. `keyring` usa ese backend de forma
nativa.

- **Aislamiento estricto:** la UI muestra **un perfil activo a la vez**;
  cambiar de perfil limpia empresa, departamentos y datos cargados. Nunca se
  fusionan. Esto implementa el "no se deben ver ni mezclar" de Svei.

#### 1.a [R2] Memoria de última carpeta de guardado (por perfil)

**Qué:** al generar un reporte, recordar la carpeta donde el usuario guardó por
última vez y sugerirla como default la próxima vez — igual que ya se hace con
el nombre de archivo dinámico por fecha, pero **persistente entre sesiones y
por perfil** (la carpeta de la empresa A no se usa para la empresa B; respeta
el aislamiento).

- **Dónde vive:** campo `ultima_carpeta_guardado` dentro del perfil, en
  `profiles.json`. No es secreto, así que va en el JSON, no en `keyring`.
- **Cuándo se escribe:** después de una exportación exitosa, se guarda la
  carpeta (`os.path.dirname`) del archivo generado en el perfil activo.
- **Comportamiento si la carpeta ya no existe** (unidad de red desconectada,
  USB removido, carpeta borrada): **nunca tronar**. Al abrir el diálogo de
  guardar, validar con `os.path.isdir`; si no existe, degradar en cascada a:
  (1) la carpeta de Documentos del usuario, (2) el Escritorio, (3) el
  directorio de trabajo actual. Registrar el evento en el log ("La carpeta
  recordada ya no está disponible; usando Documentos"), sin bloquear ni
  mostrar error modal.

**Corrección de bug ligada (memoria fantasma):** el código actual APARENTA
recordar servidor y usuario — los lee con `self.config.get("server")`
(`contpaqi_exporter.py:541`) y `self.config.get("db_user")` (`:570`) — pero
**es una funcionalidad rota: nada llama a `save_config()` con esos datos**;
`save_config` solo se invoca desde `save_known_password` (`:110`), que guarda
la contraseña. Es decir, hoy se persiste lo único que NO debería (la
contraseña) y no se persiste lo que sí debería (servidor, usuario, carpeta).
La Fase 1 **corrige esto de raíz**: al pasar al modelo de perfiles, se persisten
de verdad `servidor`, `usuario_sql`, filtros y `ultima_carpeta_guardado` en el
perfil; la contraseña sale del JSON hacia `keyring`. No tratar la memoria de
carpeta como feature nueva desde cero, sino como parte de arreglar esta
persistencia a medias.

#### 1.b [R2] Almacenamiento de contraseña — profundidad operativa

**Fallback si `keyring` no encuentra backend seguro.** En Windows con `keyring`
instalado, el backend por defecto es el Windows Credential Manager
(`WinVaultKeyring`) y normalmente está disponible. Si por alguna razón no hay
backend seguro (servicio deshabilitado, entorno atípico):

- **REGLA DURA: jamás caer en silencio a texto plano.** Prohibido cualquier
  archivo de respaldo con la contraseña en claro.
- Detectar la ausencia de backend seguro **al arranque** (consultar
  `keyring.get_keyring()` / capturar el error del backend).
- Avisar al usuario de forma clara (banner/log) y operar en **"modo sin
  recordar contraseña"**: se pide la contraseña cada sesión, se mantiene solo
  en memoria mientras la app está abierta, y nunca se escribe a disco.

**Flujo de UX (captura, guardado y confirmación sin revelar):**

1. **Primera vez por perfil:** el usuario teclea la contraseña y prueba
   conexión. Solo **cuando la conexión es exitosa**, la app guarda la
   contraseña en `keyring` bajo `contpaqi_reporteador/<profile_id>` (no antes:
   no tiene caso guardar una clave que no sirvió).
2. **Sesiones siguientes:** el campo de contraseña arranca **vacío y
   enmascarado**; junto a él, un indicador **"contraseña guardada ✔"**. Al
   conectar, la app lee la clave desde `keyring` a memoria **sin mostrarla**
   nunca en el input por defecto. El botón "Mostrar" solo revela lo que el
   usuario teclee manualmente en esa sesión.
3. **Cambiar contraseña:** si el usuario teclea una nueva y conecta con éxito,
   se sobrescribe la de `keyring`. Botón explícito "Olvidar contraseña de este
   perfil" que borra la entrada de `keyring`.

#### 1.c [R2] "Auto-Buscar" contraseña default de ContpaqI — SE CONSERVA

Reencuadre respecto a la revisión 1 (ver 1.4 punto 2 para el porqué). **No se
elimina.** Es una comodidad legítima para el contador que administra el
servidor pero no recibió la contraseña default del instalador.

- **Se conserva** el botón que prueba contraseñas contra el servidor
  seleccionado por el usuario.
- **Se recorta el diccionario** a las contraseñas default **reales de
  ContpaqI** (p. ej. `Compac1`, `COMPAC1`, `compac`, `Contpaqi1`, `contpaqi`,
  `CONTPAQI1`, y "en blanco"). Se quitan las genéricas ajenas a ContpaqI
  (`123456`, `admin`, `Admin123`, `root`, `1234`, `sql`, `server`, `master`,
  `temporal`) — no aportan y hacen que parezca herramienta de ataque.
- **Al acertar, el resultado se guarda en `keyring`** (flujo 1.b), no en texto
  plano. Se eliminan `save_known_password`/`load_known_passwords` que escribían
  a `config.json`; el diccionario recortado queda como constante de código.
- **Encuadre en UI:** etiqueta que deje claro que prueba las contraseñas
  *default* de ContpaqI contra el servidor que el usuario eligió (recuperación
  de acceso legítima), no un escáner genérico.

### Decisión 2 (Fase 2) — Nómina es el foco; contabilidad queda separada, con costura para el futuro

**Qué:** NO conectar el sistema de nómina con contabilidad/facturación ahora.
El valor inmediato es el motor de reportes de nómina personalizados.

**Por qué:** replicar la Balanza/EdoR desde SQL directo arriesga no cuadrar
con el reporte oficial (lógica de cuentas especiales, tipo de cambio). Y
`cfdi-app` / `importar_xml_a_contabilidad` son codebases separadas (fuera de
alcance). **Recomendación:** diseñar el registro de reportes (Sección 5) de
modo que una categoría "Contabilidad" pueda enchufarse **después** por el
mismo mecanismo — dejar la costura, no construirla. Para el antojo puntual de
Balanza/EdoR: un simple botón "abrir en ContpaqI" (lanzador, opción b) es un
agregado barato a futuro, no ahora.

### Decisión 3 (Fase 3) — Motor de reportes: modelo "dataset base + vistas + presets"

**Qué:** la clave de escalabilidad. En lugar de codear 35 consultas sueltas,
tres capas:

1. **Datasets base** (pocos, bien diseñados): fuentes de datos crudas —
   Empleados, Movimientos de nómina del periodo, Acumulados anuales,
   Incidencias, Préstamos/Descuentos, etc. Cada uno es una consulta
   parametrizada sobre el esquema introspectado.
2. **Vistas / variantes** (infinitas, sin código nuevo): selección de
   columnas + filtros + agrupación + orden sobre un dataset base. Aquí viven
   las "minucias" (con/sin SDI, por departamento, días laborados): son
   configuración, no reportes nuevos.
3. **Presets** (guardables por el usuario): una vista con nombre que Svei
   guarda y reutiliza ("Nómina quincenal ECA sin SDI por depto").

Encima de eso, un **catálogo de reportes predefinidos** (Sección 4) que son
presets de fábrica listos para usar. Este modelo responde directo a la
frustración de Svei: las variantes son vistas; los reportes de verdad son
datasets nuevos.

### Decisión 4 (Fase 4) — Migrar a PySide6 (secuenciado, no de golpe)

**Qué:** reescribir la GUI en **PySide6**, pero **después** de extraer el
motor.

**Por qué PySide6:**
- Consistencia con `cfdi-app`, que ya usa PySide6 (misma suite de Svei).
- Light/dark nativo vía QSS + paletas (lo que Tkinter no da).
- El Tkinter actual es un monolito imperativo de ~600 líneas que **no** va a
  sostener un registro dinámico de 35 reportes, filtros generados al vuelo,
  previsualización de temas y un diseñador de plantilla de recibos. El
  modelo/vista de Qt (`QAbstractTableModel`, `QTableView`), el theming por
  QSS y la riqueza de widgets son necesarios para el motor.

**Cómo (secuencia obligatoria):**
1. Primero extraer la capa de datos + motor de reportes a módulos
   **agnósticos de UI** (Sección 5). El Tkinter actual sigue funcionando
   sobre esos módulos durante la transición.
2. Luego construir la UI nueva en PySide6 sobre el mismo motor.
3. Retirar el Tkinter cuando la UI Qt cubra el flujo completo.

Migrar se justifica precisamente porque vamos a escalar el número de
reportes; hacerlo sobre Tkinter costaría más a mediano plazo.

### Decisión 5 (Fase 5) — Sistema de temas + marca blanca + fórmula de tint

**Qué:** un sistema de 8 temas para las tablas de reporte (Sección 8), cada
uno con tipografía; un **slot de marca blanca** (logo, color primario,
tipografía) para que un despacho use la app como marca propia; y una
**función de suavizado de color (tint)** documentada (Sección 7) porque Excel
no soporta transparencia real en rellenos.

**Por qué 8 y con marca blanca:** Svei ve potencial de venta a despachos que
usan ContpaqI. El costo marginal de varios temas es bajo (son datos, no
código) y el slot de marca blanca es lo que vuelve la app vendible. Se invierte
en la arquitectura del branding, no en pulir cada tema a mano.

---

## 4. Catálogo de reportes propuesto (~35)

> Estos son **datasets base** (D) o **reportes predefinidos** (presets de
> fábrica, P). Las variantes de cada uno (con/sin columnas, agrupaciones) NO
> se listan aparte: son vistas del mismo dataset. Los nombres de tabla/columna
> exactos se **descubren en implementación** por introspección
> (`INFORMATION_SCHEMA`), extendiendo `get_real_col`; ContpaqI Nóminas usa
> tanto códigos (`nom10001`, `nom10005`…) como nombres (`NomEmpleado`,
> `NomMovimiento`…) según versión.

### A. Plantilla / directorio (dataset base: Empleados)
1. **Directorio de empleados** (el actual) — datos generales. `[D+P]`
2. **Altas y bajas del periodo** — movimientos de personal (ingreso/baja). `[P]`
3. **Antigüedad y aniversarios** — fecha de ingreso, antigüedad calculada;
   alimenta prima de antigüedad, vacaciones y prima vacacional. `[P]`
4. **Cumpleaños del mes** — RH. `[P]`
5. **Plantilla por departamento/puesto** — headcount, organigrama tabular. `[P]`
6. **Empleados con datos incompletos/ inválidos** — RFC/CURP/NSS faltante o
   mal formado. Auditoría de calidad de datos **antes de timbrar CFDI de
   nómina**. Alto valor. `[P]`
7. **Vencimiento de contratos** — para eventuales/temporales. `[P]`

### B. Percepciones y deducciones (dataset base: Movimientos del periodo)
8. **Pre-nómina / nómina del periodo** — todas las percepciones y deducciones
   por empleado. `[D+P]`
9. **Concentrado de nómina por departamento** — totales por área. `[P]`
10. **Comparativo periodo vs. periodo** — variaciones (detecta errores de
    captura antes de pagar). `[P]`
11. **Horas extra** — dobles/triples por empleado y periodo. `[P]`
12. **Faltas / incidencias / ausentismo** — con impacto en pago. `[P]`
13. **Comisiones y bonos** — conceptos variables. `[P]`
14. **Préstamos y descuentos** — saldos de préstamos, INFONACOT, pensión
    alimenticia. `[D+P]`
15. **Dispersión bancaria** — layout para banco: CLABE + neto por empleado. `[P]`

### C. Obligaciones fiscales (IMSS, INFONAVIT, ISR, ISN)
16. **Base de cotización IMSS / SDI por empleado** — cruce con SUA. `[D+P]`
17. **Cuotas obrero-patronales IMSS estimadas** — por empleado y patronal. `[P]`
18. **Retenciones ISR + subsidio al empleo** por periodo. `[P]`
19. **INFONAVIT** — créditos y factor de descuento por empleado. `[P]`
20. **Impuesto Sobre Nómina estatal (ISN Jalisco)** — base gravable estatal.
    Svei está en Zapopan; ISN Jalisco aplica. `[P]`
21. **Conciliación CFDI de nómina timbrados vs. nómina interna** — detecta
    recibos no timbrados o con diferencias. Alto valor de cumplimiento. `[P]`

### D. Provisiones y prestaciones (pasivo laboral)
22. **Aguinaldo** — provisión y cálculo (días de ley o superiores). `[P]`
23. **Prima vacacional** — devengada por aniversario. `[P]`
24. **Vacaciones** — días generados/gozados/pendientes; saldo = pasivo. `[P]`
25. **PTU** — base, reparto, tope (3 meses de salario o promedio del último
    año, el que beneficie). `[P]`
26. **Prima de antigüedad** — pasivo laboral con tope de 2 SMG. `[P]`
27. **Provisión de pasivo laboral total** — consolida 22–26; es el puente
    natural hacia contabilidad (costura de la Decisión 2). `[P]`

### E. Liquidación / término de relación laboral
28. **Finiquito** — partes proporcionales de aguinaldo, vacaciones y prima
    vacacional. `[P]`
29. **Liquidación / indemnización** — finiquito + 3 meses + 20 días por año +
    prima de antigüedad. `[P]`
30. **Constancias** — carta de trabajo y constancia de percepciones. `[P]`

### F. Anuales / declaraciones
31. **Nómina anual por empleado** — acumulados para la anual de ISR. `[D+P]`
32. **Constancia de sueldos y salarios / retenciones anuales.** `[P]`

### G. Generador de recibos personalizable (idea estrella — Sección 6)
33. **Recibos de nómina personalizados** — trae todas las cantidades y arma
    una plantilla imprimible con logo + colores institucionales, más rápida y
    flexible que el "sobre recibo" de ContpaqI. `[P]`

### H. Fuente externa (subcontratados fuera de ContpaqI — Sección 6)
34. **Consolidado ContpaqI + externos** — mezcla empleados de ContpaqI con
    subcontratados capturados/importados, aplicando los mismos factores. `[P]`
35. **Reporte 100% personalizable** — el usuario define columnas, fórmulas y
    fuente; el caso más general del motor de la Decisión 3. `[P]`

> **Nota de implementación:** los datasets B, C, D, F requieren tocar tablas
> de **movimientos y acumulados** de ContpaqI Nóminas, no solo `nom10001`.
> Primer paso de la Fase 3: mapear por introspección las tablas de conceptos
> (percepciones/deducciones), movimientos por periodo, periodos y acumulados
> anuales, y registrar los mapeos lógicos→físicos en el módulo de
> `mapeo-esquema` (Sección 5). No inventar nombres de tabla de memoria:
> descubrirlos con `INFORMATION_SCHEMA` sobre una BD real de muestra.

---

## 5. Arquitectura de módulos objetivo

Estructura de carpetas propuesta (kebab-case en carpetas, snake_case en
Python, por la Sección 11):

```
contpaqi-reporteador/
├── acceso-datos/
│   ├── conexion.py            # construir cadena, conectar, drivers ODBC
│   ├── introspeccion.py       # descubrir tablas/columnas por versión
│   └── ejecutar_consulta.py   # correr SQL parametrizado → DataFrame
├── mapeo-esquema/
│   └── mapeo_columnas.py      # nombre lógico → columna física (formaliza get_real_col)
├── reportes/
│   ├── registro.py            # catálogo: enumera reportes disponibles
│   ├── contrato.py            # interfaz que todo reporte cumple
│   ├── datasets/              # un módulo por dataset base
│   │   ├── empleados.py
│   │   ├── movimientos.py
│   │   └── acumulados.py
│   └── predefinidos/          # presets de fábrica (Sección 4)
├── exportacion/
│   ├── motor_excel.py         # openpyxl: escribe y estiliza
│   ├── temas.py               # las 8 variantes (Sección 8)
│   ├── tint.py                # función de suavizado de color (Sección 7)
│   └── plantilla_recibos.py   # generador de recibos (Sección 6)
├── datos-externos/
│   └── subcontratados.py      # captura/importación de trabajadores externos
├── config/
│   ├── perfiles.py            # leer/escribir profiles.json (sin secretos)
│   └── credenciales.py        # keyring / Windows Credential Manager
├── ui/                        # PySide6 (Fase 4); delgada, solo orquesta
├── profiles.json              # NO versionar
├── requirements.txt           # pyodbc, pandas, openpyxl, keyring, PySide6
└── HANDOFF.md                 # este documento
```

**Contrato de reporte** (`reportes/contrato.py`) — cada reporte se
autodescribe para que la UI lo enumere y genere sus filtros al vuelo. Campos
del contrato (descriptivo, no código):

- `id` y `nombre_visible`, `categoria` (A–H).
- `dataset_base` que consume.
- `filtros_soportados` (p. ej. solo activos, departamento, periodo, rango de
  fechas) → la UI los dibuja dinámicamente.
- `columnas_disponibles` y `columnas_por_defecto` → la capa de vistas.
- `construir_consulta(contexto)` → SQL parametrizado o DataFrame.
- `postprocesar(df)` → cálculos derivados (antigüedad, proporcionales).
- `tema_por_defecto`.

La UI **no conoce reportes específicos**: recorre el `registro` y renderiza.
Agregar un reporte = agregar un módulo al registro, sin tocar la UI.

**Regla de oro de seguridad SQL:** los valores (IDs, fechas) van siempre
**parametrizados** (`?` de pyodbc), nunca por f-string. Los nombres de
tabla/columna que se deban interpolar (por el esquema dinámico) se validan
antes contra el whitelist devuelto por la introspección — nunca se aceptan
directo de entrada del usuario.

### 5.1 [R2] Mapeo del monolito → módulos (divide and conquer)

Para que la sesión de implementación no tenga que re-decidir dónde va cada
trozo del `contpaqi_exporter.py` actual (1,034 líneas), este es el destino
propuesto función por función. Recordar la regla dura de comentarios
(Sección 11): al mover una función, sus comentarios viajan con ella.

| Trozo actual (contpaqi_exporter.py) | Líneas | Destino |
|---|---|---|
| `ODBC_DRIVERS_PRIORITY`, `get_installed_sql_drivers` | 31–42 | `acceso-datos/conexion.py` |
| `load_config`, `save_config` | 71–81 | `config/perfiles.py` (reemplazado por el modelo de perfiles) |
| `DEFAULT_PASSWORDS` (recortado, ver 1.c) | 84–89 | `config/credenciales.py` (constante) |
| `load_known_passwords`, `save_known_password` | 91–110 | **Se eliminan** (escribían a `config.json`); el guardado va a `keyring` en `config/credenciales.py` |
| SDK: `SDK_INSTALL_PATHS`, `detect_contpaqi_path`, `sdk_test_connection` | 115–141 | **Eliminar** (feature muerta, hallazgo 1.7.a) o dejar stub documentado |
| `_net_use_hosts`, `detect_sql_servers` | 146–174 | `acceso-datos/conexion.py` (detección de servidores) |
| `_build_cs`, `test_connection` | 179–211 | `acceso-datos/conexion.py` |
| `get_databases` | 213–221 | `acceso-datos/introspeccion.py` |
| `get_departments` | 223–251 | `acceso-datos/introspeccion.py` + catálogo en `reportes/datasets/` |
| `safe_str` (hoy código muerto) | 256–258 | `exportacion/` como utilidad, si se decide usar; si no, eliminar |
| `get_real_col` + descubrimiento de columnas dentro de `export_data` | 280–352 | `mapeo-esquema/mapeo_columnas.py` (formaliza el patrón) |
| Armado de query de empleados en `export_data` | 260–376 | `reportes/datasets/empleados.py` (implementa el contrato) |
| Estilos de Excel dentro de `export_data` (`BLUE`, `ROW_ALT`, bordes, totales, autofilter) | 378–434 | `exportacion/motor_excel.py` + `temas.py` + `tint.py` |
| Clase `App` completa (GUI Tkinter, estilos, log, threads) | 439–1029 | `ui/` (se reescribe en PySide6, Fase 4). Durante la transición sigue en Tkinter pero **llamando a los módulos nuevos**, no con lógica embebida |
| Métodos de threading (`_auto_check_passwords`, `_start_scan`, `_load_databases`, `_load_departments`, `_execute`) | 860–1009 | `ui/` orquesta; la lógica de datos que hoy contienen se muda a `acceso-datos`/`reportes` y la UI solo la invoca |
| `if __name__ == "__main__"` | 1031–1034 | Entry point (`main.py` o `ui/app.py`) |

> Nota: las líneas son del commit auditado `58aca0f`; verificar contra el
> estado real antes de mover, por si hubo cambios posteriores.

---

## 6. Generador de recibos personalizable + subcontratados

Dos piezas que Svei pidió explícitamente y que son diferenciador comercial:

### 6.1 Plantilla de recibos (`exportacion/plantilla_recibos.py`)

- Toma el dataset de movimientos del periodo (todas las percepciones y
  deducciones por empleado) y genera un documento imprimible por empleado.
- **Personalizable:** logo del despacho/empresa, colores institucionales
  (usa el sistema de temas y la función tint de la Sección 7), tipografía.
- Objetivo: **más rápido y flexible que el "sobre recibo" de ContpaqI**.
- Formato de salida: primero Excel (una hoja/bloque por recibo, listo para
  imprimir); dejar la costura para PDF más adelante.

### 6.2 Subcontratados / fuente externa (`datos-externos/subcontratados.py`)

- Permite **capturar o importar** (desde Excel) trabajadores que NO están en
  ContpaqI, con los mismos campos/factores.
- El dataset "Consolidado ContpaqI + externos" (reporte 34) une ambas fuentes
  bajo un esquema común, para que los reportes y recibos los traten igual.
- Persistencia de los externos: por perfil, en el mismo modelo aislado de la
  Decisión 1 (no se mezclan entre empresas).

---

## 7. Fórmula de suavizado de color (tint) — explícita

**Problema:** Excel no soporta transparencia real en el relleno de celdas. Si
el usuario captura un color institucional fuerte (p. ej. `#99763c` o
`#3d425f`) y se usa a saturación completa para el bandeado de filas, cansa la
vista y compite con el texto. **Solución:** mezclar el color con blanco en un
porcentaje (concepto de "tint" de teoría del color) para generar
automáticamente una versión más clara.

### 7.1 Definición

`tint(color_hex, p)` donde `p` ∈ [0, 1] es la fracción de mezcla hacia blanco
(`p = 0` deja el color igual; `p = 1` lo vuelve blanco puro):

```
Para cada canal C ∈ {R, G, B}:
    C_nuevo = round( C_original + (255 - C_original) * p )

Pasos:
1. Quitar el '#'. Parsear RRGGBB a tres enteros R, G, B (base 16).
2. Aplicar la fórmula por canal.
3. Recomponer a hex de 2 dígitos por canal (con cero a la izquierda si hace falta).
4. Devolver '#RRGGBB' en mayúsculas (o sin '#' si openpyxl lo pide así).
```

### 7.2 Valores recomendados

- **Fila alternada (banding):** `p = 0.88` (muy claro, descansado).
- **Fila de totales / realce suave:** `p = 0.70`.
- **Encabezado:** color a **fuerza completa** (`p = 0`), con texto blanco.

### 7.3 Ejemplos verificados (con los colores que dio Svei)

`tint("#99763c", 0.88)`:
```
R = 153 + (255-153)*0.88 = 153 + 89.76 = 242.76 → 243 → F3
G = 118 + (255-118)*0.88 = 118 + 120.56 = 238.56 → 239 → EF
B =  60 + (255- 60)*0.88 =  60 +171.60 = 231.60 → 232 → E8
Resultado: #F3EFE8  (beige cálido suave)
```

`tint("#3d425f", 0.88)`:
```
R =  61 + (255- 61)*0.88 =  61 +170.72 = 231.72 → 232 → E8
G =  66 + (255- 66)*0.88 =  66 +166.32 = 232.32 → 232 → E8
B =  95 + (255- 95)*0.88 =  95 +140.80 = 235.80 → 236 → EC
Resultado: #E8E8EC  (lavanda-gris frío suave)
```

Estos dos colores (`#99763c` dorado y `#3d425f` azul-gris) deben quedar como
el **tema corporativo por defecto** (encabezado a fuerza completa, banding con
el tint 0.88 de arriba). Son la paleta institucional de Svei; ver también su
memoria de marca.

---

## 8. Sistema de temas (8 variantes, con tipografía)

Cada tema define: relleno de encabezado, color de fuente del encabezado,
banding de filas (color base + `p` de tint), color/estilo de borde, estilo de
la fila de totales, tipografía (fuente de encabezado, fuente de cuerpo,
tamaños) y formato numérico. Todos derivan los tonos claros con la función de
la Sección 7 — nunca se hardcodean tonos claros a mano.

1. **Corporativo Neutro (default).** Encabezado en color primario de marca
   (`#3d425f`), banding tint 0.88, totales tint 0.70. Tipografía: Calibri /
   Aptos. Es el punto de partida institucional.
2. **Alto Contraste.** Encabezado azul muy oscuro/negro, banding gris claro
   marcado, bordes definidos. Legible en impresión b/n y para accesibilidad.
   Tipografía: Arial.
3. **Escala de Grises.** Todo monocromático; banding en grises. Para impresión
   económica. Tipografía: Arial.
4. **Minimalista sin bandeado.** Sin relleno de filas; solo líneas
   horizontales finas y borde inferior grueso en el encabezado. Máxima
   sobriedad. Tipografía: Aptos / Calibri Light.
5. **Financiero / Contable.** Números tabulares alineados, verde/rojo para
   signos, totales fuertemente resaltados; estética de estado financiero.
   Tipografía: Cambria (serif) para cifras.
6. **Cálido Institucional (arena/dorado).** Basado en `#99763c`; encabezado
   dorado, banding beige (tint 0.88 → `#F3EFE8`). Tipografía: Georgia /
   Calibri.
7. **Frío Institucional (azul-gris).** Basado en `#3d425f`; banding
   lavanda-gris (tint 0.88 → `#E8E8EC`). Tipografía: Calibri.
8. **Marca Blanca (personalizable) — el slot vendible.** Toma del despacho:
   logo, color primario y tipografía. Deriva automáticamente encabezado
   (primario a fuerza completa), banding (tint 0.88 del primario) y totales
   (tint 0.70). Este es el tema que vuelve la app marca-blanca sin tocar
   código.

**Slot de branding (transversal a toda la app):** logo + color primario +
tipografía se guardan por perfil/instalación. La plantilla de recibos
(Sección 6) y todos los temas los consumen. Mantener este slot presente en
todo el diseño desde el inicio, aunque hoy solo lo use el despacho de Svei.

---

## 9. Roadmap por fases (con criterios de salida)

> Rama por fase; probar local antes de merge a `main` (Sección 12).

**Fase 0 — Auditoría base.** ✅ *Completada por este documento.*
Salida: hallazgos con evidencia (Sección 1), README confirmado, roadmap
creado (esta sección).

**Fase 0.5 — Higiene de seguridad y repo (BLOQUEANTE, hacer primero).**
Acciones de la Sección 10.
Criterio de salida: `.gitignore` en su lugar; `config.json` y `*.xlsx` fuera
de tracking; contraseña ya no en texto plano en el árbol de trabajo;
`requirements.txt` creado.

**Fase 1 — Memoria de entorno y credenciales (Decisión 1).**
Perfiles multi-empresa aislados; `profiles.json` sin secretos; contraseña en
Windows Credential Manager vía `keyring` con fallback seguro (1.b); memoria de
última carpeta de guardado por perfil (1.a) y corrección del bug de memoria
fantasma; **conservar** el "Auto-Buscar" reencuadrado y con diccionario
recortado que ahora guarda a `keyring` (1.c).
Criterio de salida: se puede crear ≥2 perfiles, cambiar entre ellos con
limpieza total de estado, y reconectar sin re-teclear la contraseña; ninguna
contraseña se escribe en disco en claro; al reabrir, la carpeta de guardado
sugerida es la última usada por ese perfil (o un fallback si ya no existe, sin
tronar).

**Fase 2 — Extracción del motor (Decisión 3, parte 1) + costura contable
(Decisión 2).**
Separar `acceso-datos`, `mapeo-esquema`, `reportes` (contrato + registro +
datasets base) del código de UI. El Tkinter actual sigue funcionando sobre
los módulos nuevos.
Criterio de salida: el reporte de empleados actual corre íntegro a través del
nuevo motor (mismos resultados), con el registro enumerando al menos ese
reporte; cero regresiones.

**Fase 3 — Catálogo de reportes (Decisión 3, parte 2).**
Mapear por introspección las tablas de movimientos/conceptos/periodos/
acumulados; implementar datasets base B–F; construir los predefinidos de la
Sección 4 por prioridad de Svei (empezar por: pre-nómina del periodo,
concentrado por departamento, datos incompletos antes de timbrar, y recibos).
Criterio de salida: ≥10 reportes predefinidos funcionando end-to-end sobre una
BD real; agregar un reporte nuevo no requiere tocar la UI.

**Fase 4 — Migración a PySide6 (Decisión 4).**
UI nueva sobre el motor; enumera el registro y genera filtros dinámicos;
retirar Tkinter.
Criterio de salida: la UI Qt cubre el flujo completo (perfiles → empresa →
reporte → filtros → exportar) con light/dark nativo; Tkinter eliminado.

**Fase 5 — Temas, marca blanca y recibos (Decisión 5 + Sección 6).**
Implementar `tint.py`, los 8 temas, el slot de branding y el generador de
recibos.
Criterio de salida: el usuario elige tema por reporte; carga logo/color/fuente
de marca blanca y se refleja en tablas y recibos; la función tint produce los
valores verificados de la Sección 7.3.

**Fase 6 (futuro, no diseñar a detalle aún) — Contabilidad / lanzador
ContpaqI.** Enchufar categoría contable o botón "abrir en ContpaqI" por la
costura de la Decisión 2, si Svei lo pide.

---

## 10. Acciones inmediatas de seguridad y repo (Fase 0.5)

Hacer **antes** de escribir features:

1. **Crear `.gitignore`** con al menos: `config.json`, `profiles.json`,
   `*.xlsx`, `__pycache__/`, `.venv/`, `*.pyc`.
2. **Sacar `config.json` de tracking:** `git rm --cached config.json`
   (conservar el archivo local si se necesita como referencia, pero ya no
   versionarlo).
3. **Sacar los `*.xlsx` con PII** del repo (moverlos fuera o dejarlos solo
   local ya ignorados). No deben poder subirse con un `git add .`.
4. **Sobre la contraseña commiteada (`Compac1` en `4620e7c`):** es la
   contraseña *default pública* de ContpaqI, así que la sensibilidad real es
   baja y **no** amerita reescribir el historial de golpe. Pero: (a) dejar de
   versionar `config.json` (pasos 1–2), y (b) **si la app se va a distribuir o
   el repo se hará público**, ahí sí purgar el historial (`git filter-repo` o
   BFG) y rotar cualquier contraseña real que se haya usado.
5. **Crear `requirements.txt`** con dependencias pinneadas:
   `pyodbc`, `pandas`, `openpyxl`, `keyring`, y (desde Fase 4) `PySide6`.
6. **[R2] Reencuadrar (no eliminar) el "Auto-Buscar"** como parte de la
   Fase 1: conservar el botón, recortar el diccionario a contraseñas default
   reales de ContpaqI, y hacer que el acierto se guarde en `keyring` en vez de
   texto plano. Eliminar sí `save_known_password`/`load_known_passwords` (la
   parte que escribía a `config.json`). Ver Decisión 1, sección 1.c.

---

## 11. Reglas de estilo de código (de Svei — íntegras)

- Funciones pequeñas de propósito único.
- Comentarios explicativos estilo "profesor de salón de clases" — el código
  debe poder debuggearse en una sesión fresca de IA sin contexto previo.
- Nombres en lenguaje natural, descriptivos.
- Prohibidos los patrones ingeniosos-pero-crípticos.
- Archivos chicos, arquitectura modular, separación de responsabilidades.
- Carpetas: lowercase kebab-case. Python: snake_case. Sin espacios, acentos ni
  caracteres especiales en nombres de archivos.

### [R2] Regla dura — código autoexplicativo + comentarios persistentes

Agregada por Svei; se suma a las reglas de arriba, no las diluye:

- El código en Python debe **contar su propia historia** al leerlo: los
  nombres de variables, funciones y clases describen **qué hacen y por qué**,
  no solo de qué tipo son. Si una función se entiende con solo leer su firma y
  sus nombres de variables, no necesita comentario adicional.
- Cuando el nombre **no basta para explicar el POR QUÉ** (una regla de negocio
  no obvia, una particularidad de ContpaqI, un workaround), **se exige un
  comentario breve que explique el motivo** — no lo que hace la línea, sino por
  qué existe.
- **REGLA DURA: los comentarios existentes NO se borran**, salvo que (1) se
  elimine la funcionalidad que describen, o (2) se reemplacen por una
  explicación mejor y más precisa. Nunca se borra un comentario por "limpieza"
  o refactor cosmético.
- **Al dividir el monolito** `contpaqi_exporter.py` en los módulos de la
  Sección 5, los comentarios existentes **viajan con la función** a su nuevo
  archivo, y se **añaden** comentarios nuevos donde el módulo destino
  introduzca contexto que antes era implícito por estar todo en un solo archivo
  (ver tabla de mapeo en 5.1).
- **Hacia adelante (PySide6, Fase 4):** la UI nueva sigue esta misma regla,
  incluyendo nombres descriptivos para señales/slots de Qt (p. ej.
  `perfil_cambiado`, `al_generar_reporte`, no `sig1`/`on_click`).

---

## 12. Git workflow (de Svei)

- Rama por fase. Probar local antes de merge a `main`.
- Deploy solo desde `main`, si aplica distribución a otros despachos.

---

*Fin del handoff. Cualquier decisión que este documento no cubra, resolverla
con el criterio de las Secciones 3 y 11, y anotarla como nueva decisión aquí
mismo.*
