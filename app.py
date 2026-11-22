import streamlit as st
from pathlib import Path
import subprocess
import glob
import sys

st.title("Banco Ganadero")

pdf = st.file_uploader(
    "Subir PDF",
    type=["pdf"]
)

if pdf:

    carpeta_temp = Path("temp")
    carpeta_temp.mkdir(exist_ok=True)

    ruta_pdf = carpeta_temp / pdf.name

    with open(ruta_pdf, "wb") as f:
        f.write(pdf.getbuffer())

    st.success("PDF guardado correctamente")

    st.write("Ruta:")
    st.write(str(ruta_pdf))

    if st.button("Procesar PDF"):

        st.write("Ejecutando agente...")

        resultado = subprocess.run(
            [
                sys.executable,
                "Agente_Banco/pdf_a_excel_web.py",
                str(ruta_pdf)
            ],
             capture_output=True,
            text=True
        )

        st.success("Agente ejecutado")

        st.write("STDOUT")
        st.text(resultado.stdout)

        st.write("STDERR")
        st.text(resultado.stderr)

        excels = sorted(
            glob.glob("Agente_Banco/Banco_Ganadero_*.xlsx"),
           reverse=True
        )

        if excels:

            ultimo_excel = excels[0]

            with open(ultimo_excel, "rb") as f:

               st.download_button(
                   label="Descargar Excel",
                   data=f,
                   file_name=Path(ultimo_excel).name,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )