import os
import glob
import shutil
from datetime import datetime
from dbfread import DBF
import oracledb

# ==========================================================
# CONFIGURACIÓN DE ORACLE Y RUTAS
# ==========================================================
ORACLE_USER = "ttm_admin"
ORACLE_PASS = ""
ORACLE_HOST = "localhost"
ORACLE_PORT = 1521
ORACLE_SERVICE = "xe"  # 'XEPDB1', 'xe' u 'orcl'

CARPETA_LOCAL_DEFECTO = r"C:\Users\feder\TTM_Repuestos_Finanzas\JULIA"
RUTA_RED_SERVIDOR = r"\\SERVIDOR_TTM\FoxPro\DATOS"
CARPETA_STAGING_RED = "./STAGING_SERVER"
BATCH_SIZE = 5000


def obtener_conexion_oracle():
    dsn = f"{ORACLE_HOST}:{ORACLE_PORT}/{ORACLE_SERVICE}"
    return oracledb.connect(user=ORACLE_USER, password=ORACLE_PASS, dsn=dsn)

def buscar_archivo_dbf(carpeta, nombre_tabla):
    carpeta_norm = os.path.normpath(carpeta)
    nombre_obj = nombre_tabla.strip().upper()
    for f in os.listdir(carpeta_norm):
        base, ext = os.path.splitext(f)
        if base.upper() == nombre_obj and ext.upper() == ".DBF":
            return os.path.join(carpeta_norm, f)
    raise FileNotFoundError(f"No se encontró '{nombre_tabla}.dbf' en: {carpeta_norm}")

def sincronizar_todo_a_oracle(carpeta_origen):
    print(f"\n🚀 Sincronizando Catálogo, Compras y Ventas desde: {carpeta_origen}")
    conn = obtener_conexion_oracle()
    cursor = conn.cursor()

    try:
        # 1. VENDEDORES (VENDED.DBF)
        ruta_ven = buscar_archivo_dbf(carpeta_origen, "VENDED")
        print("📖 Cargando VENDED.DBF...")
        cursor.execute("TRUNCATE TABLE vendedores")
        dict_ven = {}
        for r in DBF(ruta_ven, encoding='latin1', ignore_missing_memofile=True):
            cod = str(r['CODVEN']).strip()
            nom = str(r['NOMBRE']).strip()
            es_fer = "FERNANDA" in nom.upper() or cod.upper() in ["FER", "FD", "FND"]
            dict_ven[cod] = (cod, nom, 0.0 if es_fer else 0.02, 0.0025 if es_fer else 0.02, 1 if es_fer else 0)
        cursor.executemany("INSERT INTO vendedores VALUES (:1, :2, :3, :4, :5)", list(dict_ven.values()))
        conn.commit()

        # 2. CATÁLOGO DE PRODUCTOS (PRODUC.DBF)
        ruta_prod = buscar_archivo_dbf(carpeta_origen, "PRODUC")
        print("📖 Cargando PRODUC.DBF (Catálogo Maestro de Costos)...")
        cursor.execute("TRUNCATE TABLE productos")
        dict_prod = {}
        for r in DBF(ruta_prod, encoding='latin1', ignore_missing_memofile=True):
            cod = str(r['CODIGO']).strip()
            if cod:
                f_ultco = r['FEULTCO'].strftime('%Y-%m-%d') if r['FEULTCO'] else None
                dict_prod[cod] = (
                    cod,
                    str(r['DESCRIP']).strip() if r['DESCRIP'] else None,
                    float(r['PRCOSTO']) if r['PRCOSTO'] is not None else 0.0,
                    float(r['PRVENTA']) if r['PRVENTA'] is not None else 0.0,
                    float(r['PRULTCO']) if r['PRULTCO'] is not None else 0.0,
                    f_ultco,
                    str(r['FAMILIA']).strip() if r['FAMILIA'] else None
                )
        sql_prod = """
            INSERT INTO productos (cod_producto, descripcion, precio_costo, precio_venta, precio_ult_compra, fecha_ult_compra, familia)
            VALUES (:1, :2, :3, :4, :5, TO_DATE(:6, 'YYYY-MM-DD'), :7)
        """
        cursor.executemany(sql_prod, list(dict_prod.values()))
        conn.commit()
        print(f"✅ {len(dict_prod):,} productos migrados con sus costos.")

        # 3. HISTORIAL DE FACTURAS DE COMPRA (TIFACC.DBF y DTFACC.DBF)
        ruta_tifacc = buscar_archivo_dbf(carpeta_origen, "TIFACC")
        ruta_dtfacc = buscar_archivo_dbf(carpeta_origen, "DTFACC")
        print("📖 Cargando TIFACC.DBF y DTFACC.DBF (Histórico de Compras)...")
        
        cursor.execute("TRUNCATE TABLE detalle_compra")
        cursor.execute("DELETE FROM documentos_compra")
        conn.commit()

        docs_comp = {}
        for r in DBF(ruta_tifacc, encoding='latin1', ignore_missing_memofile=True):
            tipo, num = str(r['TIPO']).strip(), str(r['NUMERO']).strip()
            f_str = r['FECHA'].strftime('%Y-%m-%d') if r['FECHA'] else None
            if f_str and num:
                docs_comp[(tipo, num)] = (
                    tipo, num, f_str, 
                    str(r['RUTPROV']).strip() if r['RUTPROV'] else None,
                    float(r['NETO']) if r['NETO'] is not None else 0.0,
                    float(r['IVA']) if r['IVA'] is not None else 0.0,
                    str(r['ESTADO']).strip() if r['ESTADO'] else None
                )
        sql_comp = "INSERT INTO documentos_compra VALUES (:1, :2, TO_DATE(:3, 'YYYY-MM-DD'), :4, :5, :6, :7)"
        cursor.executemany(sql_comp, list(docs_comp.values()))

        det_comp = []
        for r in DBF(ruta_dtfacc, encoding='latin1', ignore_missing_memofile=True):
            tipo, num = str(r['TIPO']).strip(), str(r['NUMERO']).strip()
            f_str = r['FECHA'].strftime('%Y-%m-%d') if r['FECHA'] else None
            if f_str and num:
                det_comp.append((
                    tipo, num, f_str,
                    str(r['RUTPROV']).strip() if r['RUTPROV'] else None,
                    str(r['COD_ART']).strip(),
                    float(r['CANTID']) if r['CANTID'] is not None else 0.0,
                    float(r['PRECIO']) if r['PRECIO'] is not None else 0.0
                ))
        sql_detcomp = """
            INSERT INTO detalle_compra (tipo_doc, numero_doc, fecha, rut_proveedor, cod_producto, cantidad, precio_costo)
            VALUES (:1, :2, TO_DATE(:3, 'YYYY-MM-DD'), :4, :5, :6, :7)
        """
        for i in range(0, len(det_comp), BATCH_SIZE):
            cursor.executemany(sql_detcomp, det_comp[i:i + BATCH_SIZE])
        conn.commit()
        print(f"✅ {len(det_comp):,} líneas de compras históricas cargadas.")

        # 4. HISTORIAL DE VENTAS (TIFACV.DBF y DTFACV.DBF)
        ruta_tif = buscar_archivo_dbf(carpeta_origen, "TIFACV")
        ruta_dtf = buscar_archivo_dbf(carpeta_origen, "DTFACV")
        print("📖 Cargando TIFACV.DBF y DTFACV.DBF (Histórico de Ventas)...")

        cursor.execute("TRUNCATE TABLE detalle_venta")
        cursor.execute("DELETE FROM documentos_venta")
        conn.commit()

        docs_unicos = {}
        for r in DBF(ruta_tif, encoding='latin1', ignore_missing_memofile=True):
            tipo, num = str(r['TIPO']).strip(), str(r['NUMERO']).strip()
            if tipo in ['33', '39', '61']:
                f_str = r['FECHA'].strftime('%Y-%m-%d') if r['FECHA'] else None
                if f_str and num:
                    docs_unicos[(tipo, num)] = (
                        tipo, num, f_str,
                        str(r['RUTCLIE']).strip() if r['RUTCLIE'] else None,
                        str(r['CODVEN']).strip() if r['CODVEN'] else None,
                        float(r['PDESCT1']) if r['PDESCT1'] is not None else 0.0,
                        float(r['DESCTO1']) if r['DESCTO1'] is not None else 0.0,
                        float(r['NETO']) if r['NETO'] is not None else 0.0,
                        float(r['IVA']) if r['IVA'] is not None else 0.0,
                        str(r['ESTADO']).strip() if r['ESTADO'] else None
                    )
        sql_doc = """
            INSERT INTO documentos_venta (tipo_doc, numero_doc, fecha, rut_cliente, cod_vendedor, pdesct1, descto1, neto, iva, estado)
            VALUES (:1, :2, TO_DATE(:3, 'YYYY-MM-DD'), :4, :5, :6, :7, :8, :9, :10)
        """
        lista_docs = list(docs_unicos.values())
        for i in range(0, len(lista_docs), BATCH_SIZE):
            cursor.executemany(sql_doc, lista_docs[i:i + BATCH_SIZE])

        detalles_v = []
        for r in DBF(ruta_dtf, encoding='latin1', ignore_missing_memofile=True):
            tipo, num = str(r['TIPO']).strip(), str(r['NUMERO']).strip()
            if tipo in ['33', '39', '61']:
                f_str = r['FECHA'].strftime('%Y-%m-%d') if r['FECHA'] else None
                if f_str and num:
                    detalles_v.append((
                        tipo, num, f_str,
                        str(r['COD_ART']).strip(),
                        float(r['CANTID']) if r['CANTID'] is not None else 0.0,
                        float(r['PRECIO']) if r['PRECIO'] is not None else 0.0,
                        float(r['DESCTO']) if r['DESCTO'] is not None else 0.0
                    ))
        sql_det = """
            INSERT INTO detalle_venta (tipo_doc, numero_doc, fecha, cod_producto, cantidad, precio, descto)
            VALUES (:1, :2, TO_DATE(:3, 'YYYY-MM-DD'), :4, :5, :6, :7)
        """
        for i in range(0, len(detalles_v), BATCH_SIZE):
            cursor.executemany(sql_det, detalles_v[i:i + BATCH_SIZE])

        conn.commit()
        print(f"✅ {len(detalles_v):,} líneas de venta cargadas.")
        print("\n🎉 ¡Puente histórico y catálogo sincronizados al 100% en Oracle!")

    except Exception as e:
        conn.rollback()
        print(f"\n❌ Error durante la migración: {e}")
    finally:
        cursor.close()
        conn.close()

if __name__ == "__main__":
    sincronizar_todo_a_oracle(CARPETA_LOCAL_DEFECTO)