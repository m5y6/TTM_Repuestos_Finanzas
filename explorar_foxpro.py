import os
import glob
import re
from collections import deque
from dbfread import DBF
import pandas as pd

CARPETA_DATOS = "./JULIA"
ARCHIVO_ESTRUCTURA = "estructura_columnas_foxpro.xlsx"
ARCHIVO_ULTIMOS_REGISTROS = "ultimos_5_registros_foxpro.xlsx"

def listar_tablas(carpeta):
    """Encuentra todos los archivos .dbf sin duplicar nombres en Windows."""
    encontrados = glob.glob(os.path.join(carpeta, "*.dbf")) + glob.glob(os.path.join(carpeta, "*.DBF"))
    unicos = {}
    for ruta in encontrados:
        nombre_normalizado = os.path.basename(ruta).upper()
        if nombre_normalizado not in unicos:
            unicos[nombre_normalizado] = ruta
    return sorted(list(unicos.values()))

def limpiar_nombre_hoja(nombre):
    """Excel no permite nombres de hoja con más de 31 caracteres ni caracteres especiales."""
    nombre_base = os.path.splitext(nombre)[0]
    nombre_limpio = re.sub(r'[\\/*?:\[\]]', '_', nombre_base)
    return nombre_limpio[:31]

def procesar_archivos_foxpro():
    tablas = listar_tablas(CARPETA_DATOS)
    if not tablas:
        print("❌ No se encontraron archivos .dbf en la carpeta.")
        return

    print(f"📦 Se procesarán {len(tablas)} tablas únicas.\n")

    resumen_tablas = []
    detalle_columnas = []
    
    # Preparamos el escritor para el segundo Excel (Últimos 5 registros)
    with pd.ExcelWriter(ARCHIVO_ULTIMOS_REGISTROS, engine="openpyxl") as writer_registros:
        
        for idx, ruta in enumerate(tablas, 1):
            nombre_archivo = os.path.basename(ruta)
            nombre_hoja = limpiar_nombre_hoja(nombre_archivo)
            print(f"[{idx}/{len(tablas)}] Procesando: {nombre_archivo}...")

            try:
                # Lectura de la estructura y datos con codificación estándar FoxPro
                tabla = DBF(ruta, encoding='latin1', ignore_missing_memofile=True)
                total_registros = len(tabla)
                
                # 1. Extracción de estructura y nombres de columnas
                nombres_cols = [campo.name for campo in tabla.fields]
                
                resumen_tablas.append({
                    "Tabla": nombre_archivo,
                    "Total_Registros": total_registros,
                    "Total_Columnas": len(tabla.fields),
                    "Nombres_Todas_Las_Columnas": ", ".join(nombres_cols)
                })

                for pos, campo in enumerate(tabla.fields, 1):
                    detalle_columnas.append({
                        "Tabla": nombre_archivo,
                        "Nro_Columna": pos,
                        "Nombre_Columna": campo.name,
                        "Tipo_Dato": campo.type,
                        "Longitud": campo.length
                    })

                # 2. Extracción de los ÚLTIMOS 5 registros usando deque (eficiente en memoria)
                ultimos_5 = deque(tabla, maxlen=5)
                
                if ultimos_5:
                    df_ultimos = pd.DataFrame(list(ultimos_5))
                else:
                    # Si la tabla está vacía, guarda las columnas con mensaje
                    df_ultimos = pd.DataFrame(columns=nombres_cols if nombres_cols else ["Sin_Columnas"])
                    df_ultimos.loc[0] = ["(Tabla sin registros)"] * len(df_ultimos.columns)

                # Escribir hoja en el Excel de registros
                df_ultimos.to_excel(writer_registros, sheet_name=nombre_hoja, index=False)

            except Exception as e:
                print(f"⚠️ Error al leer {nombre_archivo}: {e}")
                resumen_tablas.append({
                    "Tabla": nombre_archivo,
                    "Total_Registros": "Error",
                    "Total_Columnas": "Error",
                    "Nombres_Todas_Las_Columnas": str(e)
                })

    # Guardar el primer Excel con la estructura completa
    print("\n⏳ Guardando archivo de estructura...")
    with pd.ExcelWriter(ARCHIVO_ESTRUCTURA, engine="openpyxl") as writer_estructura:
        pd.DataFrame(resumen_tablas).to_excel(writer_estructura, sheet_name="Resumen_General", index=False)
        pd.DataFrame(detalle_columnas).to_excel(writer_estructura, sheet_name="Detalle_Todas_Columnas", index=False)

    print(f"✅ Archivo 1 generado: {ARCHIVO_ESTRUCTURA}")
    print(f"✅ Archivo 2 generado: {ARCHIVO_ULTIMOS_REGISTROS}")

if __name__ == "__main__":
    procesar_archivos_foxpro()