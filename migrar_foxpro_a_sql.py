import os
import datetime
from dbfread import DBF
import pandas as pd
import oracledb

# ==========================================
# CONFIGURACIÓN DE CONEXIÓN ORACLE
# ==========================================
ORACLE_USER = "ttm_admin"
ORACLE_PASS = ""
ORACLE_HOST = "localhost"
ORACLE_PORT = 1521
# Ajusta el servicio si usas XE ('XEPDB1' o 'xe') o Enterprise ('orcl')
ORACLE_SERVICE = "xe" 

CARPETA_DATOS = "."

def encontrar_archivo(carpeta, nombre_sin_ext):
    carpeta_norm = os.path.normpath(carpeta)
    nombre_obj = nombre_sin_ext.strip().upper()
    for f in os.listdir(carpeta_norm):
        b, ext = os.path.splitext(f)
        if b.upper() == nombre_obj and ext.upper() == ".DBF":
            return os.path.join(carpeta_norm, f)
    raise FileNotFoundError(f"No se encontró el archivo {nombre_sin_ext}.dbf")

def migrar_a_oracle():
    print("🔌 Conectando a Oracle Database...")
    dsn = f"{ORACLE_HOST}:{ORACLE_PORT}/{ORACLE_SERVICE}"
    
    try:
        conn = oracledb.connect(user=ORACLE_USER, password=ORACLE_PASS, dsn=dsn)
        cursor = conn.cursor()
        print("✅ Conexión establecida con éxito.\n")
    except Exception as e:
        print(f"❌ Error al conectar a Oracle: {e}")
        print("💡 Tip: Si usas SID 'xe' en lugar de servicio, revisa la variable ORACLE_SERVICE.")
        return

    try:
        # 1. Migrar Vendedores (VENDED.DBF)
        ruta_vended = encontrar_archivo(CARPETA_DATOS, "VENDED")
        print(f"📖 Leyendo {os.path.basename(ruta_vended)}...")
        df_ven = pd.DataFrame(iter(DBF(ruta_vended, encoding='latin1', ignore_missing_memofile=True)))
        
        cursor.execute("TRUNCATE TABLE vendedores")
        sql_ven = """
            INSERT INTO vendedores (cod_vendedor, nombre, tasa_general, tasa_herramientas, solo_herramientas)
            VALUES (:1, :2, :3, :4, :5)
        """
        datos_ven = []
        for _, r in df_ven.iterrows():
            cod = str(r['CODVEN']).strip()
            nom = str(r['NOMBRE']).strip()
            es_fernanda = "FERNANDA" in nom.upper() or cod.upper() in ["FER", "FD", "FND"]
            
            tasa_gen = 0.0 if es_fernanda else 0.02
            tasa_herr = 0.0025 if es_fernanda else 0.02
            solo_herr = 1 if es_fernanda else 0
            datos_ven.append((cod, nom, tasa_gen, tasa_herr, solo_herr))

        cursor.executemany(sql_ven, datos_ven)
        conn.commit()
        print(f"✅ {len(datos_ven)} vendedores migrados.\n")

        # 2. Migrar Cabeceras (TIFACV.DBF)
        ruta_tifacv = encontrar_archivo(CARPETA_DATOS, "TIFACV")
        print(f"📖 Leyendo {os.path.basename(ruta_tifacv)}...")
        df_tif = pd.DataFrame(iter(DBF(ruta_tifacv, encoding='latin1', ignore_missing_memofile=True)))
        
        # Limpiar dependencias
        cursor.execute("TRUNCATE TABLE detalle_venta")
        cursor.execute("DELETE FROM documentos_venta")
        conn.commit()

        sql_doc = """
            INSERT INTO documentos_venta (
                tipo_doc, numero_doc, fecha, rut_cliente, cod_vendedor, 
                pdesct1, descto1, neto, iva, estado
            ) VALUES (:1, :2, TO_DATE(:3, 'YYYY-MM-DD'), :4, :5, :6, :7, :8, :9, :10)
        """
        datos_doc = []
        for _, r in df_tif.iterrows():
            tipo = str(r['TIPO']).strip()
            if tipo in ['33', '39', '61']:
                f_str = r['FECHA'].strftime('%Y-%m-%d') if pd.notnull(r['FECHA']) else None
                if f_str:
                    datos_doc.append((
                        tipo,
                        str(r['NUMERO']).strip(),
                        f_str,
                        str(r['RUTCLIE']).strip() if pd.notnull(r['RUTCLIE']) else None,
                        str(r['CODVEN']).strip() if pd.notnull(r['CODVEN']) else None,
                        float(r['PDESCT1']) if pd.notnull(r['PDESCT1']) else 0.0,
                        float(r['DESCTO1']) if pd.notnull(r['DESCTO1']) else 0.0,
                        float(r['NETO']) if pd.notnull(r['NETO']) else 0.0,
                        float(r['IVA']) if pd.notnull(r['IVA']) else 0.0,
                        str(r['ESTADO']).strip() if pd.notnull(r['ESTADO']) else None
                    ))

        # Inserción por lotes
        cursor.executemany(sql_doc, datos_doc)
        conn.commit()
        print(f"✅ {len(datos_doc)} documentos cabecera migrados.\n")

        # 3. Migrar Detalle (DTFACV.DBF)
        ruta_dtfacv = encontrar_archivo(CARPETA_DATOS, "DTFACV")
        print(f"📖 Leyendo {os.path.basename(ruta_dtfacv)}...")
        df_dtf = pd.DataFrame(iter(DBF(ruta_dtfacv, encoding='latin1', ignore_missing_memofile=True)))

        sql_det = """
            INSERT INTO detalle_venta (
                tipo_doc, numero_doc, fecha, cod_producto, cantidad, precio, descto
            ) VALUES (:1, :2, TO_DATE(:3, 'YYYY-MM-DD'), :4, :5, :6, :7)
        """
        datos_det = []
        for _, r in df_dtf.iterrows():
            tipo = str(r['TIPO']).strip()
            if tipo in ['33', '39', '61']:
                f_str = r['FECHA'].strftime('%Y-%m-%d') if pd.notnull(r['FECHA']) else None
                if f_str:
                    datos_det.append((
                        tipo,
                        str(r['NUMERO']).strip(),
                        f_str,
                        str(r['COD_ART']).strip(),
                        float(r['CANTID']) if pd.notnull(r['CANTID']) else 0.0,
                        float(r['PRECIO']) if pd.notnull(r['PRECIO']) else 0.0,
                        float(r['DESCTO']) if pd.notnull(r['DESCTO']) else 0.0
                    ))

        cursor.executemany(sql_det, datos_det)
        conn.commit()
        print(f"✅ {len(datos_det)} líneas de detalle migradas.\n")

        print("🎉 ¡Migración completa a Oracle finalizada exitosamente!")

    except Exception as e:
        conn.rollback()
        print(f"❌ Error durante la migración: {e}")
    finally:
        cursor.close()
        conn.close()

if __name__ == "__main__":
    migrar_a_oracle()