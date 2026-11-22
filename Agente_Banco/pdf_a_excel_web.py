import re
from pathlib import Path
from datetime import datetime
import sys

import pandas as pd
from pypdf import PdfReader


# =====================================================
# RUTAS
# =====================================================

# Este script debe estar en:
# AgenteVentas / Archivos_Agente / Banco / pdf_a_excel.py

CARPETA = Path(__file__).resolve().parent

pdfs = sorted(
    CARPETA.glob("EXT_BG*.pdf"),
    key=lambda x: x.stat().st_mtime,
    reverse=True
)


PDF_FILE = None

# Si viene una ruta por parámetro
if len(sys.argv) > 1:

    PDF_FILE = Path(sys.argv[1])

# Si no viene una ruta, seguir como antes
elif pdfs:

    PDF_FILE = pdfs[0]

else:

    raise FileNotFoundError(
        f"No se encontró ningún PDF EXT_BG*.pdf en:\n{CARPETA}"
    )

TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")

OUTPUT_FILE = (
    CARPETA /
    f"Banco_Ganadero_{TIMESTAMP}.xlsx"
)


# =====================================================
# FUNCIONES
# =====================================================

def limpiar_numero(valor):
    """
    Convierte textos monetarios del banco a número real.

    Ejemplos:
        +31.00      -> 31.00
        -50.00      -> -50.00
        1,029.12    -> 1029.12
        +7,400.00   -> 7400.00
    """

    valor = str(valor).strip()
    valor = valor.replace(",", "")
    valor = valor.replace("+", "")

    try:
        return float(valor)
    except:
        return 0.0


def limpiar_descripcion(texto):
    """
    Deja la descripción en una sola línea,
    sin saltos de línea ni espacios duplicados.
    """

    return " ".join(str(texto).split())


# =====================================================
# LEER PDF
# =====================================================

print()
print("=" * 60)
print("CONVERTIDOR BANCO GANADERO PDF A EXCEL")
print("=" * 60)
print(f"PDF encontrado:")
print(PDF_FILE)
print()

print("Leyendo PDF...")

reader = PdfReader(str(PDF_FILE))

texto = "\n".join(
    page.extract_text() or ""
    for page in reader.pages
)
texto_original = texto
    
print()
print()

# =====================================================
# LIMPIEZA DEL TEXTO
# =====================================================

texto = re.sub(
    r"Página\s+\d+\s+de\s+\d+",
    "",
    texto
)

texto = re.sub(
    r"La información del extracto está sujeta a verificación del Banco Ganadero\.",
    "",
    texto
)

texto = re.sub(
    r"SALDO ANTERIOR:\s*[\d,]+\.\d{2}",
    "",
    texto
)

# =====================================================
# NORMALIZAR MOVIMIENTOS PEGADOS
# =====================================================

texto = re.sub(
    r"(?<!\n)(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}\s+WEB/MOV/KSC)",
    r"\n\1",
    texto
)

texto = re.sub(
    r"\n+",
    "\n",
    texto
)

# =====================================================
# EXTRAER MOVIMIENTOS
# FORMATO NUEVO BANCO GANADERO
# =====================================================

filas = []

lineas = texto.splitlines()

inicio_patron = re.compile(
    r"^\d{2}/\d{2}/\d{4}\s+"
    r"\d{2}:\d{2}:\d{2}\s+"
    r"WEB/MOV/KSC\s+"
    r"\d+"
)

indices = []

for i, linea in enumerate(lineas):

    if inicio_patron.match(linea.strip()):

        indices.append(i)



for pos, inicio in enumerate(indices):

    fin = (
        indices[pos + 1]
        if pos + 1 < len(indices)
        else len(lineas)
    )

    bloque = lineas[inicio:fin]

    
    if len(bloque) == 0:        
        continue

    cabecera = bloque[0].strip()

    m = re.match(
        r"(\d{2}/\d{2}/\d{4})\s+"
        r"(\d{2}:\d{2}:\d{2})\s+"
        r"(WEB\/MOV\/KSC)\s+"
        r"(\d+)",
        cabecera
    )

    if not m:

        print()
        print("BLOQUE DESCARTADO (CABECERA)")
        print("-" * 50)
        
        print("\n".join(bloque))

        print("-" * 50)
      
        continue

    fecha = m.group(1)
    hora = m.group(2)
    oficina_canal = m.group(3)
    transaccion = m.group(4)

    resto_cabecera = cabecera[m.end():].strip()
   
    bloque_texto = resto_cabecera
    
    if len(bloque) > 1:
        bloque_texto += "\n" + "\n".join(bloque[1:])
    bloque_texto = re.sub(
      r"SALDO ACTUAL\s*:.*",
       "",
       bloque_texto,
       flags=re.S
    )
    # eliminar texto de cierre del extracto
    bloque_texto = re.sub(
       r"SALDO ACTUAL\s*:.*",
       "",
        bloque_texto,
       flags=re.S
    )

    matches = list(re.finditer(
       r"(-?\d[\d,]*\.\d{2})\s+"
       r"([+]?\d[\d,]*\.\d{2})\s+"
       r"(\d[\d,]*\.\d{2})",
       bloque_texto
    ))

    if not matches:

       print()
       print("=" * 80)
       print("BLOQUE SIN MONTOS")
       print("=" * 80)

       print("\n".join(bloque))

       print("=" * 80)

       continue

    monto_match = matches[-1]

    debito = monto_match.group(1)
    credito = monto_match.group(2)
    saldo = monto_match.group(3)

    descripcion = bloque_texto[
        :monto_match.start()
    ].strip()

    descripcion = limpiar_descripcion(
        descripcion
    )
    
    debito_num = limpiar_numero(
        debito
    )

    credito_num = limpiar_numero(
        credito
    )

    saldo_num = limpiar_numero(
        saldo
    )

    if abs(debito_num) < 0.0001:
        debito_num = None

    if abs(credito_num) < 0.0001:
        credito_num = None


    filas.append(
        [
            fecha,
            hora,
            oficina_canal,
            transaccion,
            descripcion,
            debito_num,
            credito_num,
            saldo_num
        ]
    )

print()

# =====================================================
# DATAFRAME
# =====================================================

df = pd.DataFrame(
    filas,
    columns=[
        "Fecha",
        "Hora",
        "Oficina/Canal",
        "Transacción",
        "Descripción",
        "Débito",
        "Crédito",
        "Saldo"
    ]
)

if len(df) == 0:
    raise ValueError(
        "No se encontraron movimientos en el PDF. "
        "Revisar si el formato del extracto cambió."
    )


# =====================================================
# NORMALIZAR FECHA Y HORA
# =====================================================

df["Fecha"] = pd.to_datetime(
    df["Fecha"],
    dayfirst=True,
    errors="coerce"
).dt.strftime("%Y-%m-%d")

df["Hora"] = pd.to_datetime(
    df["Hora"],
    format="%H:%M:%S",
    errors="coerce"
).dt.strftime("%H:%M:%S")

# =====================================================
# SALDO CALCULADO Y CONTROL
# =====================================================

# Extraer saldo anterior desde el PDF
m_saldo = re.search(
    r"SALDO ANTERIOR:\s*([\d,]+\.\d{2})",
    texto_original
)

saldo_inicial = 0

if m_saldo:
    saldo_inicial = limpiar_numero(
        m_saldo.group(1)
    )

saldo_calculado = saldo_inicial

saldos_calculados = []
controles = []

for _, fila in df.iterrows():

    credito = (
        fila["Crédito"]
        if pd.notna(fila["Crédito"])
        else 0
    )

    debito = (
        fila["Débito"]
        if pd.notna(fila["Débito"])
        else 0
    )

    saldo_calculado += credito + debito

    saldos_calculados.append(
        round(saldo_calculado, 2)
    )

    controles.append(
        round(
            saldo_calculado - fila["Saldo"],
            2
        )
    )

df["Saldo Calculado"] = saldos_calculados
df["Control"] = controles

# =====================================================
# EXPORTAR A EXCEL
# =====================================================

print("Generando Excel...")

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    df.to_excel(
        writer,
        sheet_name="Movimientos",
        index=False
    )

    ws = writer.sheets["Movimientos"]

    # Anchos de columnas

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 100
    ws.column_dimensions["F"].width = 15
    ws.column_dimensions["G"].width = 15
    ws.column_dimensions["H"].width = 15
    ws.column_dimensions["I"].width = 18
    ws.column_dimensions["J"].width = 15

    # Formato numérico para Débito, Crédito y Saldo

    for col in ["F", "G", "H"]:
        for cell in ws[col][1:]:
            cell.number_format = "#,##0.00"

    # Filtro y encabezado congelado

    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"


# =====================================================
# RESULTADO
# =====================================================

total_creditos = df["Crédito"].fillna(0).sum()
total_debitos = df["Débito"].fillna(0).sum()

print()
print("=" * 60)
print("PROCESO FINALIZADO")
print("=" * 60)
print(f"Movimientos encontrados : {len(df)}")
print(f"Total créditos          : {total_creditos:,.2f}")
print(f"Total débitos           : {total_debitos:,.2f}")
print()
print("Archivo generado:")
print(OUTPUT_FILE)
print("=" * 60)
print()