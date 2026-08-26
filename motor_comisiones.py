import os
import datetime
from dbfread import DBF
import pandas as pd

# ==========================================
# CONFIGURACIÓN Y PARÁMETROS DE NEGOCIO
# ==========================================
CARPETA_DATOS = "."
TASA_GENERAL = 0.02          # 2.0% para todos los vendedores
TASA_FERNANDA_HERR = 0.0025   # 0.25% exclusivo de Fernanda en herramientas

def calcular_comision_linea(cod_art, cod_ven, nombre_ven, monto_neto_linea):
    """
    Aplica las reglas de negocio por producto y vendedor:
    - TTM Service (Mano de Obra 'SER_'): 0%
    - Fernanda: 0.25% exclusivamente si es herramienta ('HER-'), 0% en otros productos.
    - Demás vendedores: 2% general.
    """
    cod_art_upper = str(cod_art).strip().upper()
    nombre_ven_upper = str(nombre_ven).strip().upper()
    cod_ven_upper = str(cod_ven).strip().upper()

    # 1. Regla TTM Service: Mano de Obra no comisiona
    if cod_art_upper.startswith("SER_"):
        return 0.0, 0.0, "Servicio (Mano de Obra - Sin Comisión)"

    # 2. Regla Exclusiva para Fernanda
    es_fernanda = "FERNANDA" in nombre_ven_upper or cod_ven_upper in ["FER", "FD", "FND"]
    es_herramienta = cod_art_upper.startswith("HER-") or cod_art_upper.startswith("HER")

    if es_fernanda:
        if es_herramienta:
            tasa = TASA_FERNANDA_HERR
            categoria = "Herramientas (Comisión Fernanda 0.25%)"
        else:
            tasa = 0.0
            categoria = "No comisionable (Fernanda no comisiona repuestos)"
    else:
        # 3. Regla General para el resto de vendedores
        tasa = TASA_GENERAL
        categoria = "Repuestos / Venta General (2.0%)"

    monto_comision = round(monto_neto_linea * tasa)
    return tasa, monto_comision, categoria


def cargar_tabla_dbf(nombre_archivo, encoding='latin1'):
    """Carga un archivo DBF en un DataFrame de pandas de forma segura."""
    ruta = os.path.join(CARPETA_DATOS, nombre_archivo)
    if not os.path.exists(ruta):
        # Intentar en minúsculas si falla
        ruta = os.path.join(CARPETA_DATOS, nombre_archivo.lower())
    
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"No se encontró el archivo: {nombre_archivo}")
    
    tabla = DBF(ruta, encoding=encoding, ignore_missing_memofile=True)
    return pd.DataFrame(iter(tabla))


def procesar_comisiones(fecha_inicio, fecha_fin, archivo_salida="Liquidacion_Comisiones.xlsx"):
    """
    Ejecuta el cálculo mensual de comisiones cruzando ventas, devoluciones y vendedores.
    """
    print(f"🚀 Iniciando cálculo de comisiones desde {fecha_inicio} hasta {fecha_fin}...")

    # 1. Cargar tablas maestras
    df_tifacv = cargar_tabla_dbf("TIFACV.DBF")
    df_dtfacv = cargar_tabla_dbf("DTFACV.DBF")
    df_vended = cargar_tabla_dbf("VENDED.DBF")

    # Limpieza de espacios en códigos
    df_tifacv['TIPO'] = df_tifacv['TIPO'].astype(str).str.strip()
    df_tifacv['NUMERO'] = df_tifacv['NUMERO'].astype(str).str.strip()
    df_tifacv['CODVEN'] = df_tifacv['CODVEN'].astype(str).str.strip()
    
    df_dtfacv['TIPO'] = df_dtfacv['TIPO'].astype(str).str.strip()
    df_dtfacv['NUMERO'] = df_dtfacv['NUMERO'].astype(str).str.strip()
    df_dtfacv['COD_ART'] = df_dtfacv['COD_ART'].astype(str).str.strip()
    
    df_vended['CODVEN'] = df_vended['CODVEN'].astype(str).str.strip()
    df_vended['NOMBRE'] = df_vended['NOMBRE'].astype(str).str.strip()

    # 2. Filtrar documentos del período y tipos válidos (33: Factura, 39: Boleta, 61: Nota Crédito)
    df_tifacv['FECHA'] = pd.to_datetime(df_tifacv['FECHA'])
    
    mask_periodo = (df_tifacv['FECHA'] >= pd.to_datetime(fecha_inicio)) & \
                   (df_tifacv['FECHA'] <= pd.to_datetime(fecha_fin))
    mask_tipos = df_tifacv['TIPO'].isin(['33', '39', '61'])
    
    cabeceras_filtradas = df_tifacv[mask_periodo & mask_tipos].copy()

    # 3. Cruzar con catálogo de vendedores
    cabeceras_con_vendedor = cabeceras_filtradas.merge(
        df_vended[['CODVEN', 'NOMBRE']], 
        on='CODVEN', 
        how='left'
    )
    cabeceras_con_vendedor['NOMBRE'] = cabeceras_con_vendedor['NOMBRE'].fillna('SIN VENDEDOR ASIGNADO')

    # 4. Cruzar detalle de productos con cabecera
    detalle_completo = df_dtfacv.merge(
        cabeceras_con_vendedor[['TIPO', 'NUMERO', 'FECHA', 'CODVEN', 'NOMBRE', 'PDESCT1', 'RUTCLIE']],
        on=['TIPO', 'NUMERO'],
        how='inner'
    )

    # 5. Aplicar Reglas de Negocio
    alertas_log = []
    registros_procesados = []

    for _, fila in detalle_completo.iterrows():
        tipo_doc = fila['TIPO']
        cod_art = fila['COD_ART']
        cod_ven = fila['CODVEN']
        nombre_ven = fila['NOMBRE']
        cant = float(fila['CANTID']) if pd.notnull(fila['CANTID']) else 0.0
        precio = float(fila['PRECIO']) if pd.notnull(fila['PRECIO']) else 0.0
        descto_linea = float(fila['DESCTO']) if pd.notnull(fila['DESCTO']) else 0.0
        descto_global = float(fila['PDESCT1']) if pd.notnull(fila['PDESCT1']) else 0.0

        # Factor multiplicador (+1 venta, -1 devolución/NC)
        signo = -1 if tipo_doc == '61' else 1
        
        # Cálculo de monto neto de la línea con descuentos aplicados
        subtotal = cant * precio
        monto_neto_linea = subtotal * (1 - (descto_linea / 100)) * (1 - (descto_global / 100)) * signo

        # Cálculo centralizado de comisión por línea
        tasa_aplicada, monto_comision, categoria_item = calcular_comision_linea(
            cod_art=cod_art,
            cod_ven=cod_ven,
            nombre_ven=nombre_ven,
            monto_neto_linea=monto_neto_linea
        )

        # Registro de alerta preventiva si falta vendedor
        if nombre_ven == 'SIN VENDEDOR ASIGNADO':
            alertas_log.append({
                "TIPO_DOC": tipo_doc,
                "NUMERO_DOC": fila['NUMERO'],
                "FECHA": fila['FECHA'].strftime('%Y-%m-%d'),
                "MOTIVO": "Documento emitido sin código de vendedor en FoxPro"
            })

        registros_procesados.append({
            "FECHA_EMISION": fila['FECHA'].strftime('%Y-%m-%d'),
            "TIPO_DOC": "FACTURA" if tipo_doc == '33' else ("BOLETA" if tipo_doc == '39' else "NOTA CRÉDITO"),
            "NUMERO_DOC": fila['NUMERO'],
            "COD_VENDEDOR": cod_ven,
            "NOMBRE_VENDEDOR": nombre_ven,
            "RUT_CLIENTE": fila['RUTCLIE_y'] if 'RUTCLIE_y' in fila else fila['RUTCLIE'],
            "COD_PRODUCTO": cod_art,
            "CATEGORIA": categoria_item,
            "CANTIDAD": cant,
            "PRECIO_UNIT": precio,
            "DESCTO_GLOBAL_%": descto_global,
            "NETO_LINEA": round(monto_neto_linea),
            "TASA_COMISION_%": tasa_aplicada * 100,
            "MONTO_COMISION": round(monto_comision)
        })

    df_resultado = pd.DataFrame(registros_procesados)

    # 6. Tabla Resumen Agregada por Vendedor
    resumen_vendedores = df_resultado.groupby(['COD_VENDEDOR', 'NOMBRE_VENDEDOR']).agg(
        TOTAL_VENTAS_NETAS=('NETO_LINEA', lambda x: x[x > 0].sum()),
        TOTAL_DEVOLUCIONES_NC=('NETO_LINEA', lambda x: abs(x[x < 0].sum())),
        NETO_COMISIONABLE=('NETO_LINEA', 'sum'),
        TOTAL_COMISION_A_PAGAR=('MONTO_COMISION', lambda x: max(0, x.sum()))  # Topado a 0 si da negativo
    ).reset_index()

    # 7. Exportación a Excel consolidado
    with pd.ExcelWriter(archivo_salida, engine="openpyxl") as writer:
        resumen_vendedores.to_excel(writer, sheet_name="Resumen_Vendedores", index=False)
        df_resultado.to_excel(writer, sheet_name="Detalle_Comisiones", index=False)
        
        if alertas_log:
            pd.DataFrame(alertas_log).drop_duplicates().to_excel(writer, sheet_name="Alertas_Revision", index=False)

    print(f"✅ Proceso finalizado. Archivo generado: '{archivo_salida}'")
    return resumen_vendedores


if __name__ == "__main__":
    # Período de corte contable (del 26 del mes anterior al 25 del mes actual)
    FECHA_CORTE_INICIO = "2026-07-26"
    FECHA_CORTE_FIN = "2026-08-25"
    
    resumen = procesar_comisiones(FECHA_CORTE_INICIO, FECHA_CORTE_FIN)
    print("\n--- RESUMEN DE COMISIONES POR VENDEDOR ---")
    print(resumen.to_string(index=False))