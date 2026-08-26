import os
import warnings
import datetime
import oracledb
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Silenciar advertencias internas de conexión DBAPI2 de pandas
warnings.filterwarnings('ignore', category=UserWarning)

# ==========================================================
# 1. CONFIGURACIÓN DE CONEXIÓN A ORACLE
# ==========================================================
ORACLE_USER = "ttm_admin"
ORACLE_PASS = ""
ORACLE_HOST = "localhost"
ORACLE_PORT = 1521
ORACLE_SERVICE = "xe"  # 'XEPDB1', 'xe' u 'orcl'

def get_oracle_connection():
    dsn = f"{ORACLE_HOST}:{ORACLE_PORT}/{ORACLE_SERVICE}"
    return oracledb.connect(user=ORACLE_USER, password=ORACLE_PASS, dsn=dsn)

def asegurar_tablas_oracle():
    """Crea la tabla de ajustes si aún no existe en el esquema."""
    conn = get_oracle_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            CREATE TABLE comisiones_linea_ajustes (
                tipo_doc     VARCHAR2(5) NOT NULL,
                numero_doc   VARCHAR2(20) NOT NULL,
                cod_producto VARCHAR2(30) NOT NULL,
                cod_vendedor VARCHAR2(10) NOT NULL,
                porc_comis   NUMBER(6,2) NOT NULL,
                CONSTRAINT pk_comis_ajustes PRIMARY KEY (tipo_doc, numero_doc, cod_producto, cod_vendedor)
            )
        """)
        conn.commit()
    except oracledb.DatabaseError:
        pass
    finally:
        cursor.close()
        conn.close()


# ==========================================================
# 2. FUNCIONES DE VALIDACIÓN Y CÁLCULO (ÁMBITO GLOBAL)
# ==========================================================
def validar_y_normalizar_rango_fechas(f_ini_str, f_fin_str):
    """
    Valida que los inputs contengan fechas reales y completas.
    Normaliza formatos comunes a YYYY-MM-DD.
    """
    def parse_fecha(txt, nombre_campo):
        valor = txt.strip()
        if not valor:
            return None, f"El campo '{nombre_campo}' no puede estar vacío."
        
        formatos_admitidos = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"]
        for fmt in formatos_admitidos:
            try:
                dt = datetime.datetime.strptime(valor, fmt).date()
                return dt, None
            except ValueError:
                pass

        return None, (
            f"El valor '{valor}' en '{nombre_campo}' no es una fecha válida.\n\n"
            f"Por favor ingrese la fecha completa en formato AAAA-MM-DD.\n"
            f"Ejemplo: 2026-08-25"
        )

    dt_ini, err_ini = parse_fecha(f_ini_str, "Fecha Inicio")
    if err_ini:
        return None, None, err_ini

    dt_fin, err_fin = parse_fecha(f_fin_str, "Fecha Fin")
    if err_fin:
        return None, None, err_fin

    if dt_ini > dt_fin:
        return None, None, "La 'Fecha Inicio' no puede ser posterior a la 'Fecha Fin'."

    return dt_ini.strftime("%Y-%m-%d"), dt_fin.strftime("%Y-%m-%d"), None


def obtener_resumen_y_detalle(fecha_inicio, fecha_fin):
    query = """
    SELECT 
        d.tipo_doc AS tipo_doc_raw,
        CASE 
            WHEN d.tipo_doc = '33' THEN 'Factura'
            WHEN d.tipo_doc = '39' THEN 'Boleta'
            WHEN d.tipo_doc = '61' THEN 'Nota Crédito'
            ELSE d.tipo_doc 
        END AS tipo_doc,
        d.numero_doc AS numero,
        TO_CHAR(d.fecha, 'YYYY-MM-DD') AS fecha_doc,
        CASE 
            WHEN UPPER(TRIM(doc.estado)) = 'CANCELA' THEN 'CANCELADO'
            ELSE 'CRÉDITO (NO PAGADO)'
        END AS estado_doc,
        d.cod_producto AS producto,
        NVL(p.descripcion, 'SIN DESCRIPCIÓN') AS descripcion_producto,
        (d.cantidad * (CASE WHEN d.tipo_doc = '61' THEN -1.0 ELSE 1.0 END)) AS cant,
        d.precio AS precio_orig,
        NVL(p.precio_costo, 0.0) AS costo_unitario,
        ROUND((NVL(d.descto, 0) + NVL(doc.pdesct1, 0)), 2) AS desc_porc,
        ROUND(d.precio * (1 - (NVL(d.descto, 0)/100.0)) * (1 - (NVL(doc.pdesct1, 0)/100.0)), 2) AS precio_final,
        ROUND((d.cantidad * d.precio * (1 - (NVL(d.descto, 0)/100.0)) * (1 - (NVL(doc.pdesct1, 0)/100.0))) * 
        (CASE WHEN d.tipo_doc = '61' THEN -1.0 ELSE 1.0 END), 2) AS tot_venta,
        ROUND((d.cantidad * NVL(p.precio_costo, 0.0)) * 
        (CASE WHEN d.tipo_doc = '61' THEN -1.0 ELSE 1.0 END), 2) AS tot_costo,
        ROUND(((d.cantidad * d.precio * (1 - (NVL(d.descto, 0)/100.0)) * (1 - (NVL(doc.pdesct1, 0)/100.0))) - 
               (d.cantidad * NVL(p.precio_costo, 0.0))) * 
              (CASE WHEN d.tipo_doc = '61' THEN -1.0 ELSE 1.0 END), 2) AS ganancia_total,
        CASE 
            WHEN UPPER(d.cod_producto) LIKE 'HER-%' OR UPPER(d.cod_producto) LIKE 'HER%' THEN 'FERNANDA DUGAN'
            ELSE NVL(v.nombre, 'SIN VENDEDOR ASIGNADO')
        END AS empleado,
        CASE 
            WHEN UPPER(d.cod_producto) LIKE 'HER-%' OR UPPER(d.cod_producto) LIKE 'HER%' THEN '131'
            ELSE NVL(doc.cod_vendedor, 'SIN_COD')
        END AS cod_vendedor,
        NVL(aj.porc_comis, 
            CASE 
                WHEN UPPER(d.cod_producto) LIKE 'SER_%' THEN 0.0
                WHEN UPPER(d.cod_producto) LIKE 'HER-%' OR UPPER(d.cod_producto) LIKE 'HER%' THEN 0.25
                ELSE (NVL(v.tasa_general, 0.02) * 100.0)
            END
        ) AS porc_comis,
        CASE 
            WHEN UPPER(d.cod_producto) LIKE 'HER-%' OR UPPER(d.cod_producto) LIKE 'HER%' THEN 1 
            ELSE 0 
        END AS es_herramienta
    FROM detalle_venta d
    INNER JOIN documentos_venta doc 
        ON d.tipo_doc = doc.tipo_doc AND d.numero_doc = doc.numero_doc
    LEFT JOIN productos p 
        ON d.cod_producto = p.cod_producto
    LEFT JOIN vendedores v 
        ON doc.cod_vendedor = v.cod_vendedor
    LEFT JOIN comisiones_linea_ajustes aj
        ON d.tipo_doc = aj.tipo_doc 
       AND d.numero_doc = aj.numero_doc 
       AND d.cod_producto = aj.cod_producto 
       AND (CASE WHEN UPPER(d.cod_producto) LIKE 'HER-%' OR UPPER(d.cod_producto) LIKE 'HER%' THEN '131' ELSE NVL(doc.cod_vendedor, 'SIN_COD') END) = aj.cod_vendedor
    WHERE d.fecha >= TO_DATE(:1, 'YYYY-MM-DD') 
      AND d.fecha <= TO_DATE(:2, 'YYYY-MM-DD')
    ORDER BY d.fecha DESC, d.numero_doc DESC
    """
    conn = get_oracle_connection()
    df_det = pd.read_sql_query(query, conn, params=(fecha_inicio, fecha_fin))
    conn.close()

    if df_det.empty:
        return pd.DataFrame(), pd.DataFrame(), []

    df_det.columns = df_det.columns.str.upper()

    def calcular_pago(row):
        if row['ES_HERRAMIENTA'] == 1:
            return round(row['TOT_VENTA'] * (row['PORC_COMIS'] / 100.0), 2)
        else:
            return round(row['GANANCIA_TOTAL'] * (row['PORC_COMIS'] / 100.0), 2)

    df_det['PAGO_EMPLEADO'] = df_det.apply(calcular_pago, axis=1)
    df_det['GANANCIA_EMPRESA'] = (df_det['GANANCIA_TOTAL'] - df_det['PAGO_EMPLEADO']).round(2)
    df_det['FECHA_CALC'] = datetime.datetime.now().strftime('%Y-%m-%d')

    alertas = df_det[df_det['EMPLEADO'] == 'SIN VENDEDOR ASIGNADO'][
        ['FECHA_DOC', 'TIPO_DOC', 'NUMERO', 'PRODUCTO', 'TOT_VENTA']
    ].drop_duplicates().to_dict('records')

    df_res = df_det.groupby(['COD_VENDEDOR', 'EMPLEADO']).agg(
        TOTAL_VENTAS_NETAS=('TOT_VENTA', lambda x: x[x > 0].sum()),
        TOTAL_NOTAS_CREDITO=('TOT_VENTA', lambda x: abs(x[x < 0].sum())),
        GANANCIA_TOTAL_MARGEN=('GANANCIA_TOTAL', 'sum'),
        TOTAL_COMISION_PAGAR=('PAGO_EMPLEADO', lambda x: max(0, x.sum())),
        GANANCIA_NETA_EMPRESA=('GANANCIA_EMPRESA', 'sum')
    ).reset_index()

    columnas_historial = [
        'TIPO_DOC', 'NUMERO', 'FECHA_DOC', 'ESTADO_DOC', 'PRODUCTO', 'DESCRIPCION_PRODUCTO',
        'CANT', 'PRECIO_ORIG', 'COSTO_UNITARIO', 'DESC_PORC', 'PRECIO_FINAL', 
        'TOT_VENTA', 'TOT_COSTO', 'GANANCIA_TOTAL', 'EMPLEADO', 
        'PORC_COMIS', 'PAGO_EMPLEADO', 'GANANCIA_EMPRESA', 'FECHA_CALC',
        'TIPO_DOC_RAW', 'COD_VENDEDOR'
    ]
    df_det_formateado = df_det[columnas_historial].copy()

    return df_res, df_det_formateado, alertas


def guardar_ajuste_oracle(tipo_doc_raw, numero_doc, cod_prod, cod_ven, porc_comis):
    conn = get_oracle_connection()
    cursor = conn.cursor()
    merge_sql = """
    MERGE INTO comisiones_linea_ajustes dst
    USING (SELECT :1 AS tipo_doc, :2 AS numero_doc, :3 AS cod_producto, :4 AS cod_vendedor, :5 AS porc_comis FROM dual) src
    ON (dst.tipo_doc = src.tipo_doc AND dst.numero_doc = src.numero_doc AND dst.cod_producto = src.cod_producto AND dst.cod_vendedor = src.cod_vendedor)
    WHEN MATCHED THEN
        UPDATE SET dst.porc_comis = src.porc_comis
    WHEN NOT MATCHED THEN
        INSERT (tipo_doc, numero_doc, cod_producto, cod_vendedor, porc_comis)
        VALUES (src.tipo_doc, src.numero_doc, src.cod_producto, src.cod_vendedor, src.porc_comis)
    """
    cursor.execute(merge_sql, (tipo_doc_raw, str(numero_doc).strip(), str(cod_prod).strip(), str(cod_ven).strip(), float(porc_comis)))
    conn.commit()
    cursor.close()
    conn.close()


# ==========================================================
# 3. INTERFAZ GRÁFICA TKINTER
# ==========================================================
class AppComisionesOracle(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TTM Repuestos - Sistema de Comisiones y Rentabilidad")
        self.geometry("1300x800")
        self.configure(bg="#f4f6f9")

        asegurar_tablas_oracle()
        self.df_resumen = None
        self.df_detalle = None
        self.alertas = []

        self.construir_header()
        self.construir_notebook()
        self.verificar_conexion_inicial()

    def obtener_periodo_defecto(self):
        hoy = datetime.date.today()
        if hoy.day >= 25:
            fin = datetime.date(hoy.year, hoy.month, 25)
            if hoy.month == 1:
                inicio = datetime.date(hoy.year - 1, 12, 25)
            else:
                inicio = datetime.date(hoy.year, hoy.month - 1, 25)
        else:
            if hoy.month == 1:
                fin = datetime.date(hoy.year - 1, 12, 25)
                inicio = datetime.date(hoy.year - 1, 11, 25)
            elif hoy.month == 2:
                fin = datetime.date(hoy.year, 1, 25)
                inicio = datetime.date(hoy.year - 1, 12, 25)
            else:
                fin = datetime.date(hoy.year, hoy.month - 1, 25)
                inicio = datetime.date(hoy.year, hoy.month - 2, 25)

        return inicio.strftime("%Y-%m-%d"), fin.strftime("%Y-%m-%d")

    def construir_header(self):
        frame_top = tk.LabelFrame(self, text=" Control de Temporada y Ciclo de Corte (25 al 25) ", font=("Segoe UI", 9, "bold"), bg="#ffffff", padx=15, pady=8)
        frame_top.pack(fill="x", padx=15, pady=8)

        f_ini, f_fin = self.obtener_periodo_defecto()

        tk.Label(frame_top, text="Fecha Inicio:", bg="#ffffff").grid(row=0, column=0, padx=5, sticky="w")
        self.entry_inicio = ttk.Entry(frame_top, width=12)
        self.entry_inicio.insert(0, f_ini)
        self.entry_inicio.grid(row=0, column=1, padx=5)

        tk.Label(frame_top, text="Fecha Fin:", bg="#ffffff").grid(row=0, column=2, padx=10, sticky="w")
        self.entry_fin = ttk.Entry(frame_top, width=12)
        self.entry_fin.insert(0, f_fin)
        self.entry_fin.grid(row=0, column=3, padx=5)

        btn_calc = tk.Button(frame_top, text="⚡ Consultar Temporada", bg="#28a745", fg="white", font=("Segoe UI", 9, "bold"), command=self.ejecutar_calculo)
        btn_calc.grid(row=0, column=4, padx=15)

        btn_edit = tk.Button(frame_top, text="✏️ Modificar % Ganancia", bg="#ffc107", fg="black", font=("Segoe UI", 9, "bold"), command=self.modal_editar_comision_seleccion)
        btn_edit.grid(row=0, column=5, padx=5)

        btn_exp = tk.Button(frame_top, text="📊 Exportar Excel", bg="#17a2b8", fg="white", font=("Segoe UI", 9, "bold"), command=self.exportar_excel)
        btn_exp.grid(row=0, column=6, padx=10)

    def construir_notebook(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=15, pady=5)

        self.tab_resumen = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_resumen, text=" 📊 Resumen por Vendedor ")
        self.crear_tab_resumen()

        self.tab_detalle = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_detalle, text=" 🔍 Detalle Comisiones (Historial Completo) ")
        self.crear_tab_detalle()

        self.lbl_estado = tk.Label(self, text="Conectando a Oracle...", bd=1, relief="sunken", anchor="w", bg="#e9ecef", font=("Segoe UI", 9))
        self.lbl_estado.pack(side="bottom", fill="x")

    def crear_tab_resumen(self):
        frame_grid = ttk.Frame(self.tab_resumen)
        frame_grid.pack(fill="both", expand=True)

        cols = ("COD", "EMPLEADO", "VENTAS_NETAS", "NOTAS_CREDITO", "GANANCIA_TOTAL", "PAGO_EMPLEADO", "GANANCIA_EMPRESA")
        self.tree_res = ttk.Treeview(frame_grid, columns=cols, show="headings")
        
        self.tree_res.heading("COD", text="Cód")
        self.tree_res.heading("EMPLEADO", text="Vendedor")
        self.tree_res.heading("VENTAS_NETAS", text="Ventas Netas ($)")
        self.tree_res.heading("NOTAS_CREDITO", text="Notas Crédito ($)")
        self.tree_res.heading("GANANCIA_TOTAL", text="Ganancia / Margen ($)")
        self.tree_res.heading("PAGO_EMPLEADO", text="Pago Comisión ($)")
        self.tree_res.heading("GANANCIA_EMPRESA", text="Ganancia Empresa ($)")

        self.tree_res.column("COD", width=60, anchor="center")
        self.tree_res.column("EMPLEADO", width=250, anchor="w")
        self.tree_res.column("VENTAS_NETAS", width=140, anchor="e")
        self.tree_res.column("NOTAS_CREDITO", width=140, anchor="e")
        self.tree_res.column("GANANCIA_TOTAL", width=150, anchor="e")
        self.tree_res.column("PAGO_EMPLEADO", width=150, anchor="e")
        self.tree_res.column("GANANCIA_EMPRESA", width=150, anchor="e")

        sb_y = ttk.Scrollbar(frame_grid, orient="vertical", command=self.tree_res.yview)
        self.tree_res.configure(yscrollcommand=sb_y.set)
        self.tree_res.pack(side="left", fill="both", expand=True)
        sb_y.pack(side="right", fill="y")

    def crear_tab_detalle(self):
        frame_filtros = tk.Frame(self.tab_detalle, bg="#ffffff", bd=1, relief="groove", padx=10, pady=6)
        frame_filtros.pack(fill="x", padx=5, pady=5)

        tk.Label(frame_filtros, text="🔍 Buscar N° Docto:", bg="#ffffff", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(5, 2))
        self.entry_busqueda_num = ttk.Entry(frame_filtros, width=15)
        self.entry_busqueda_num.pack(side="left", padx=(0, 15))
        self.entry_busqueda_num.bind("<KeyRelease>", lambda event: self.aplicar_filtros_detalle())

        tk.Label(frame_filtros, text="👤 Filtrar por Empleado:", bg="#ffffff", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(5, 2))
        self.combo_filtro_emp = ttk.Combobox(frame_filtros, width=28, state="readonly")
        self.combo_filtro_emp.set("TODOS")
        self.combo_filtro_emp.pack(side="left", padx=(0, 15))
        self.combo_filtro_emp.bind("<<ComboboxSelected>>", lambda event: self.aplicar_filtros_detalle())

        btn_limpiar = ttk.Button(frame_filtros, text="🧹 Limpiar Filtros", command=self.limpiar_filtros_detalle)
        btn_limpiar.pack(side="left", padx=5)

        self.lbl_conteo_filtro = tk.Label(frame_filtros, text="", bg="#ffffff", fg="#6c757d", font=("Segoe UI", 9, "italic"))
        self.lbl_conteo_filtro.pack(side="right", padx=10)

        frame_grid = ttk.Frame(self.tab_detalle)
        frame_grid.pack(fill="both", expand=True, padx=5, pady=(0, 5))

        self.cols_det_config = [
            ("TIPO_DOC", "Tipo Doc", 90, "center"),
            ("NUMERO", "Número", 85, "center"),
            ("FECHA_DOC", "Fecha Doc", 95, "center"),
            ("ESTADO_DOC", "Estado Doc", 145, "center"),
            ("PRODUCTO", "Producto", 110, "w"),
            ("DESCRIPCION_PRODUCTO", "Descripción", 210, "w"),
            ("CANT", "Cant", 55, "center"),
            ("PRECIO_ORIG", "Precio Orig ($)", 105, "e"),
            ("COSTO_UNITARIO", "Costo ($)", 105, "e"),
            ("DESC_PORC", "Desc %", 70, "center"),
            ("PRECIO_FINAL", "Precio Final ($)", 110, "e"),
            ("TOT_VENTA", "Tot Venta ($)", 115, "e"),
            ("TOT_COSTO", "Tot Costo ($)", 115, "e"),
            ("GANANCIA_TOTAL", "Ganancia Total ($)", 125, "e"),
            ("EMPLEADO", "Empleado", 170, "w"),
            ("PORC_COMIS", "% Comis", 80, "center"),
            ("PAGO_EMPLEADO", "Pago Empleado ($)", 125, "e"),
            ("GANANCIA_EMPRESA", "Ganancia Empresa ($)", 135, "e")
        ]

        col_ids = [c[0] for c in self.cols_det_config]
        self.tree_det = ttk.Treeview(frame_grid, columns=col_ids, show="headings", selectmode="extended")

        for col_id, titulo, ancho, alineacion in self.cols_det_config:
            self.tree_det.heading(col_id, text=titulo)
            self.tree_det.column(col_id, width=ancho, anchor=alineacion)

        sb_y = ttk.Scrollbar(frame_grid, orient="vertical", command=self.tree_det.yview)
        sb_x = ttk.Scrollbar(frame_grid, orient="horizontal", command=self.tree_det.xview)
        
        self.tree_det.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)

        self.tree_det.grid(row=0, column=0, sticky="nsew")
        sb_y.grid(row=0, column=1, sticky="ns")
        sb_x.grid(row=1, column=0, sticky="ew")

        frame_grid.grid_rowconfigure(0, weight=1)
        frame_grid.grid_columnconfigure(0, weight=1)

        self.tree_det.bind("<Double-1>", lambda event: self.modal_editar_comision_seleccion())

    def aplicar_filtros_detalle(self):
        if self.df_detalle is None or self.df_detalle.empty:
            return

        df_filtrado = self.df_detalle.copy()

        busq_num = self.entry_busqueda_num.get().strip()
        if busq_num:
            busq_clean = busq_num.lstrip('0')
            if busq_clean:
                df_filtrado = df_filtrado[
                    df_filtrado['NUMERO'].astype(str).str.lstrip('0').str.contains(busq_clean, case=False, na=False)
                ]

        emp_seleccionado = self.combo_filtro_emp.get()
        if emp_seleccionado and emp_seleccionado != "TODOS":
            df_filtrado = df_filtrado[df_filtrado['EMPLEADO'] == emp_seleccionado]

        for i in self.tree_det.get_children():
            self.tree_det.delete(i)

        for _, r in df_filtrado.head(2000).iterrows():
            self.tree_det.insert("", "end", values=(
                r['TIPO_DOC'],
                r['NUMERO'],
                r['FECHA_DOC'],
                r['ESTADO_DOC'],
                r['PRODUCTO'],
                r['DESCRIPCION_PRODUCTO'],
                f"{r['CANT']:.0f}",
                f"${r['PRECIO_ORIG']:,.0f}".replace(",", "."),
                f"${r['COSTO_UNITARIO']:,.0f}".replace(",", "."),
                f"{r['DESC_PORC']:.1f}%",
                f"${r['PRECIO_FINAL']:,.0f}".replace(",", "."),
                f"${r['TOT_VENTA']:,.0f}".replace(",", "."),
                f"${r['TOT_COSTO']:,.0f}".replace(",", "."),
                f"${r['GANANCIA_TOTAL']:,.0f}".replace(",", "."),
                r['EMPLEADO'],
                f"{r['PORC_COMIS']:.2f}%",
                f"${r['PAGO_EMPLEADO']:,.0f}".replace(",", "."),
                f"${r['GANANCIA_EMPRESA']:,.0f}".replace(",", ".")
            ))

        self.lbl_conteo_filtro.config(text=f"Mostrando {len(df_filtrado):,} de {len(self.df_detalle):,} registros")

    def limpiar_filtros_detalle(self):
        self.entry_busqueda_num.delete(0, tk.END)
        self.combo_filtro_emp.set("TODOS")
        self.aplicar_filtros_detalle()

    def verificar_conexion_inicial(self):
        try:
            conn = get_oracle_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM documentos_venta")
            total = cursor.fetchone()[0]
            cursor.close()
            conn.close()
            self.lbl_estado.config(text=f"🟢 Conectado a Oracle Database ({ORACLE_SERVICE}) | {total:,} documentos cargados.")
            self.ejecutar_calculo()
        except Exception as e:
            self.lbl_estado.config(text=f"🔴 Error de conexión: {e}")

    def modal_editar_comision_seleccion(self):
        if hasattr(self, 'win_edicion_comis') and self.win_edicion_comis is not None and self.win_edicion_comis.winfo_exists():
            self.win_edicion_comis.lift()
            self.win_edicion_comis.focus_force()
            return

        seleccionados = self.tree_det.selection()
        if not seleccionados:
            messagebox.showwarning("Atención", "Seleccione al menos una línea en la tabla para modificar su comisión.")
            return

        item_id = seleccionados[0]
        vals = self.tree_det.item(item_id, "values")
        tipo_doc, num_doc, fecha, estado, prod, desc, cant, p_orig, costo, desc_p, p_fin, tot_v, tot_c, ganancia, emp, comis_actual, pago_e, gan_emp = vals

        match = self.df_detalle[(self.df_detalle['NUMERO'] == num_doc) & (self.df_detalle['PRODUCTO'] == prod)]
        if match.empty:
            return
        fila = match.iloc[0]

        self.win_edicion_comis = tk.Toplevel(self)
        self.win_edicion_comis.title(f"Ajustar % Comisión - Doc: {num_doc}")
        self.win_edicion_comis.geometry("400x330")
        self.win_edicion_comis.resizable(False, False)
        self.win_edicion_comis.transient(self)
        self.win_edicion_comis.grab_set()

        tk.Label(self.win_edicion_comis, text=f"Documento: {tipo_doc} N° {num_doc}", font=("Segoe UI", 10, "bold")).pack(pady=5)
        tk.Label(self.win_edicion_comis, text=f"Vendedor: {emp}", font=("Segoe UI", 9)).pack()
        tk.Label(self.win_edicion_comis, text=f"Producto: {prod} ({desc})", font=("Segoe UI", 9, "italic")).pack()
        tk.Label(self.win_edicion_comis, text=f"Ganancia / Margen de esta línea: {ganancia}", font=("Segoe UI", 9, "bold"), fg="#007bff").pack(pady=5)

        tk.Label(self.win_edicion_comis, text="Nuevo % de Comisión sobre Ganancia:").pack(pady=(5, 0))
        entry_porc = ttk.Entry(self.win_edicion_comis, width=12)
        entry_porc.insert(0, comis_actual.replace("%", ""))
        entry_porc.pack()
        entry_porc.focus_set()

        var_alcance = tk.StringVar(value="LINEA")
        rb1 = ttk.Radiobutton(self.win_edicion_comis, text=f"Aplicar solo a este producto ({prod})", variable=var_alcance, value="LINEA")
        rb1.pack(anchor="w", padx=40, pady=(10, 2))
        rb2 = ttk.Radiobutton(self.win_edicion_comis, text=f"Aplicar a TODOS los productos de la {tipo_doc} N° {num_doc}", variable=var_alcance, value="DOCUMENTO")
        rb2.pack(anchor="w", padx=40, pady=2)

        def guardar():
            try:
                nuevo_pct = float(entry_porc.get().strip())
                if var_alcance.get() == "LINEA":
                    guardar_ajuste_oracle(fila['TIPO_DOC_RAW'], num_doc, prod, fila['COD_VENDEDOR'], nuevo_pct)
                else:
                    lineas_doc = self.df_detalle[self.df_detalle['NUMERO'] == num_doc]
                    for _, r in lineas_doc.iterrows():
                        guardar_ajuste_oracle(r['TIPO_DOC_RAW'], num_doc, r['PRODUCTO'], r['COD_VENDEDOR'], nuevo_pct)

                self.win_edicion_comis.destroy()
                self.ejecutar_calculo()
                messagebox.showinfo("Guardado", f"Ajuste del {nuevo_pct}% guardado permanentemente en Oracle.")
            except ValueError:
                messagebox.showerror("Error", "Ingrese un número válido (ej. 25 o 20).")

        tk.Button(self.win_edicion_comis, text="💾 Guardar en Base de Datos", bg="#28a745", fg="white", font=("Segoe UI", 9, "bold"), command=guardar).pack(pady=15)

    def ejecutar_calculo(self):
        inicio_raw = self.entry_inicio.get()
        fin_raw = self.entry_fin.get()

        inicio, fin, error_msg = validar_y_normalizar_rango_fechas(inicio_raw, fin_raw)
        if error_msg:
            messagebox.showwarning("Fecha Inválida", error_msg)
            self.lbl_estado.config(text="⚠️ Error: Ingrese fechas válidas para continuar.")
            return

        self.entry_inicio.delete(0, tk.END)
        self.entry_inicio.insert(0, inicio)
        self.entry_fin.delete(0, tk.END)
        self.entry_fin.insert(0, fin)
        
        try:
            self.lbl_estado.config(text="⏳ Consultando Oracle Database...")
            self.update_idletasks()

            df_res, df_det, alertas = obtener_resumen_y_detalle(inicio, fin)
            if df_res.empty:
                messagebox.showwarning("Sin Datos", "No se encontraron transacciones en la temporada seleccionada.")
                self.lbl_estado.config(text="Sin datos en el período.")
                return

            self.df_resumen = df_res
            self.df_detalle = df_det
            self.alertas = alertas

            empleados_unicos = ["TODOS"] + sorted(list(df_det['EMPLEADO'].dropna().unique()))
            self.combo_filtro_emp['values'] = empleados_unicos

            for i in self.tree_res.get_children():
                self.tree_res.delete(i)
            for _, r in df_res.iterrows():
                self.tree_res.insert("", "end", values=(
                    r['COD_VENDEDOR'],
                    r['EMPLEADO'],
                    f"${r['TOTAL_VENTAS_NETAS']:,.0f}".replace(",", "."),
                    f"${r['TOTAL_NOTAS_CREDITO']:,.0f}".replace(",", "."),
                    f"${r['GANANCIA_TOTAL_MARGEN']:,.0f}".replace(",", "."),
                    f"${r['TOTAL_COMISION_PAGAR']:,.0f}".replace(",", "."),
                    f"${r['GANANCIA_NETA_EMPRESA']:,.0f}".replace(",", ".")
                ))

            self.aplicar_filtros_detalle()
            self.lbl_estado.config(text=f"🟢 Temporada calculada: {len(df_res)} vendedores y {len(df_det):,} líneas. (Doble clic en una fila para modificar su % de comisión).")
        except Exception as e:
            messagebox.showerror("Error al Consultar", f"Error en Oracle: {e}")
            self.lbl_estado.config(text="Error en la consulta.")

    def exportar_excel(self):
        if self.df_resumen is None or self.df_resumen.empty:
            messagebox.showwarning("Atención", "Primero consulte los datos de la temporada.")
            return

        archivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel Files", "*.xlsx")], 
            initialfile=f"Comisiones_TTM_{self.entry_inicio.get()}_a_{self.entry_fin.get()}.xlsx"
        )
        if archivo:
            try:
                cols_export = [
                    'TIPO_DOC', 'NUMERO', 'FECHA_DOC', 'ESTADO_DOC', 'PRODUCTO', 'DESCRIPCION_PRODUCTO',
                    'CANT', 'PRECIO_ORIG', 'COSTO_UNITARIO', 'DESC_PORC', 'PRECIO_FINAL', 
                    'TOT_VENTA', 'TOT_COSTO', 'GANANCIA_TOTAL', 'EMPLEADO', 
                    'PORC_COMIS', 'PAGO_EMPLEADO', 'GANANCIA_EMPRESA', 'FECHA_CALC'
                ]
                with pd.ExcelWriter(archivo, engine="openpyxl") as writer:
                    self.df_resumen.to_excel(writer, sheet_name="Resumen_Vendedores", index=False)
                    self.df_detalle[cols_export].to_excel(writer, sheet_name="Historial", index=False)
                messagebox.showinfo("Exportación Exitosa", f"Informe guardado correctamente en:\n{archivo}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo guardar el archivo: {e}")

if __name__ == "__main__":
    app = AppComisionesOracle()
    app.mainloop()