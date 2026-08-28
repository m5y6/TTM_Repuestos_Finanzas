import os
import glob
import shutil
from datetime import datetime
from dbfread import DBF
import oracledb

# ==========================================================
# CONFIGURACIÓN DE CONEXIÓN ORACLE
# ==========================================================
ORACLE_USER = "ttm_admin"
ORACLE_PASS = ""
ORACLE_HOST = "localhost"
ORACLE_PORT = 1521
ORACLE_SERVICE = "xe"

CARPETA_LOCAL_DEFECTO = r"C:\Users\feder\TTM_Repuestos_Finanzas\JULIA"
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
    return None

def sincronizar_todo_a_oracle(carpeta_origen):
    print(f"\n🚀 Sincronizando Clientes, Catálogo, Compras y Ventas desde: {carpeta_origen}")
    conn = obtener_conexion_oracle()
    cursor = conn.cursor()

    try:
        # 1. CLIENTES / COMPRADORES (CLIENT.DBF y CLIENT1.DBF)
        print("📖 Leyendo Clientes y Compradores (CLIENT.DBF / CLIENT1.DBF)...")
        dict_clientes = {}
        for nom_dbf in ["CLIENT", "CLIENT1"]:
            ruta_cli = buscar_archivo_dbf(carpeta_origen, nom_dbf)
            if ruta_cli:
                try:
                    for r in DBF(ruta_cli, encoding='latin1', ignore_missing_memofile=True):
                        rut = str(r['RUTCLI']).strip() if r.get('RUTCLI') else ""
                        nom = str(r['NOMBRE']).strip() if r.get('NOMBRE') else ""
                        if rut and nom:
                            dict_clientes[rut] = (
                                rut,
                                nom,
                                str(r['DIRECC']).strip() if r.get('DIRECC') else None,
                                str(r['COMUNA']).strip() if r.get('COMUNA') else None,
                                str(r['CIUDAD']).strip() if r.get('CIUDAD') else None,
                                str(r['GIRO']).strip() if r.get('GIRO') else None,
                                str(r['FONO']).strip() if r.get('FONO') else None,
                                str(r['EMAIL']).strip() if r.get('EMAIL') else None
                            )
                except Exception as e:
                    print(f"⚠️ Aviso al leer {nom_dbf}: {e}")

        cursor.execute("TRUNCATE TABLE clientes")
        sql_cli = """
            INSERT INTO clientes (rut_cliente, nombre, direccion, comuna, ciudad, giro, telefono, email)
            VALUES (:1, :2, :3, :4, :5, :6, :7, :8)
        """
        lista_cli = list(dict_clientes.values())
        for i in range(0, len(lista_cli), BATCH_SIZE):
            cursor.executemany(sql_cli, lista_cli[i:i + BATCH_SIZE])
        conn.commit()
        print(f"✅ {len(lista_cli):,} clientes y compradores migrados (incluido TTM Service SpA).")

        # 2. VENDEDORES (VENDED.DBF)
        ruta_ven = buscar_archivo_dbf(carpeta_origen, "VENDED")
        if ruta_ven:
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

        # 3. CATÁLOGO DE PRODUCTOS (PRODUC.DBF)
        ruta_prod = buscar_archivo_dbf(carpeta_origen, "PRODUC")
        if ruta_prod:
            print("📖 Cargando PRODUC.DBF...")
            cursor.execute("TRUNCATE TABLE productos")
            dict_prod = {}
            for r in DBF(ruta_prod, encoding='latin1', ignore_missing_memofile=True):
                cod = str(r['CODIGO']).strip()
                if cod:
                    f_ultco = r['FEULTCO'].strftime('%Y-%m-%d') if r.get('FEULTCO') else None
                    dict_prod[cod] = (
                        cod,
                        str(r['DESCRIP']).strip() if r.get('DESCRIP') else None,
                        float(r['PRCOSTO']) if r.get('PRCOSTO') is not None else 0.0,
                        float(r['PRVENTA']) if r.get('PRVENTA') is not None else 0.0,
                        float(r['PRULTCO']) if r.get('PRULTCO') is not None else 0.0,
                        f_ultco,
                        str(r['FAMILIA']).strip() if r.get('FAMILIA') else None
                    )
            sql_prod = """
                INSERT INTO productos (cod_producto, descripcion, precio_costo, precio_venta, precio_ult_compra, fecha_ult_compra, familia)
                VALUES (:1, :2, :3, :4, :5, TO_DATE(:6, 'YYYY-MM-DD'), :7)
            """
            cursor.executemany(sql_prod, list(dict_prod.values()))
            conn.commit()

        # 4. VENTAS: CABECERA Y DETALLE (TIFACV.DBF y DTFACV.DBF)
        ruta_tif = buscar_archivo_dbf(carpeta_origen, "TIFACV")
        ruta_dtf = buscar_archivo_dbf(carpeta_origen, "DTFACV")
        if ruta_tif and ruta_dtf:
            print("📖 Cargando TIFACV.DBF y DTFACV.DBF...")
            cursor.execute("TRUNCATE TABLE detalle_venta")
            cursor.execute("DELETE FROM documentos_venta")
            conn.commit()

            docs_unicos = {}
            for r in DBF(ruta_tif, encoding='latin1', ignore_missing_memofile=True):
                tipo, num = str(r['TIPO']).strip(), str(r['NUMERO']).strip()
                if tipo in ['33', '39', '61']:
                    f_str = r['FECHA'].strftime('%Y-%m-%d') if r.get('FECHA') else None
                    if f_str and num:
                        docs_unicos[(tipo, num)] = (
                            tipo, num, f_str,
                            str(r['RUTCLIE']).strip() if r.get('RUTCLIE') else None,
                            str(r['CODVEN']).strip() if r.get('CODVEN') else None,
                            float(r['PDESCT1']) if r.get('PDESCT1') is not None else 0.0,
                            float(r['DESCTO1']) if r.get('DESCTO1') is not None else 0.0,
                            float(r['NETO']) if r.get('NETO') is not None else 0.0,
                            float(r['IVA']) if r.get('IVA') is not None else 0.0,
                            str(r['ESTADO']).strip() if r.get('ESTADO') else None
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
                    f_str = r['FECHA'].strftime('%Y-%m-%d') if r.get('FECHA') else None
                    if f_str and num:
                        detalles_v.append((
                            tipo, num, f_str,
                            str(r['COD_ART']).strip(),
                            float(r['CANTID']) if r.get('CANTID') is not None else 0.0,
                            float(r['PRECIO']) if r.get('PRECIO') is not None else 0.0,
                            float(r['DESCTO']) if r.get('DESCTO') is not None else 0.0
                        ))
            sql_det = """
                INSERT INTO detalle_venta (tipo_doc, numero_doc, fecha, cod_producto, cantidad, precio, descto)
                VALUES (:1, :2, TO_DATE(:3, 'YYYY-MM-DD'), :4, :5, :6, :7)
            """
            for i in range(0, len(detalles_v), BATCH_SIZE):
                cursor.executemany(sql_det, detalles_v[i:i + BATCH_SIZE])

            conn.commit()
            print(f"✅ {len(detalles_v):,} líneas de venta migradas.")

        print("\n🎉 ¡Sincronización completa con catálogo y clientes finalizada!")

    except Exception as e:
        conn.rollback()
        print(f"\n❌ Error durante la migración: {e}")
    finally:
        cursor.close()
        conn.close()

if __name__ == "__main__":
    sincronizar_todo_a_oracle(CARPETA_LOCAL_DEFECTO)