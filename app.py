
from __future__ import annotations

import io
import os
import zipfile
import tempfile
from pathlib import Path
from contextlib import contextmanager

import streamlit as st

import planilla_fichajes_inmac_v3 as motor

st.set_page_config(
    page_title="INMAC | Planilla de fichajes",
    page_icon="📋",
    layout="centered",
)

LOGO_PATH = Path(__file__).with_name("logo_inmac.jpg")


@contextmanager
def temp_chdir(path: Path):
    anterior = Path.cwd()
    os.chdir(path)
    try:
        yield
    finally:
        os.chdir(anterior)


def guardar_upload(uploaded_file, destino: Path) -> Path:
    ruta = destino / uploaded_file.name
    ruta.write_bytes(uploaded_file.getbuffer())
    return ruta


def resumen_diagnostico(diagnostico: list[dict]) -> list[dict]:
    return [
        {
            "Archivo": Path(item["path"]).name,
            "Score planilla": item.get("score_planilla"),
            "Score fichajes": item.get("score_fichajes"),
        }
        for item in diagnostico
    ]


def procesar_desde_streamlit(uploads: list):
    if len(uploads) < 2:
        raise ValueError("Tenés que subir la planilla base y el archivo de fichajes.")

    with tempfile.TemporaryDirectory() as tmpdir:
        carpeta = Path(tmpdir)
        rutas = [guardar_upload(u, carpeta) for u in uploads]
        archivo_planilla, archivo_fichajes, diagnostico = motor.detectar_archivos([str(p) for p in rutas])

        with temp_chdir(carpeta):
            nombres_salida = motor.procesar_archivos(
                archivo_planilla=archivo_planilla,
                archivo_fichajes=archivo_fichajes,
                descargar_en_colab=False,
            )

        salidas = [carpeta / nombre for nombre in nombres_salida]
        blobs = {archivo.name: archivo.read_bytes() for archivo in salidas}

        return {
            "archivo_planilla": Path(archivo_planilla).name,
            "archivo_fichajes": Path(archivo_fichajes).name,
            "diagnostico": resumen_diagnostico(diagnostico),
            "salidas": blobs,
        }


if LOGO_PATH.exists():
    st.image(str(LOGO_PATH), width=180)

st.title("Planilla de fichajes y horas")
st.caption("INMAC · Hemoterapia / Cubiertas · Subí la planilla base y el Original del fichador.")

with st.container(border=True):
    st.subheader("Cómo usarlo")
    st.markdown(
        "1. Subí la **planilla base** de la obra y el archivo **Original** del fichador.\n"
        "2. Tocá **Procesar archivos**.\n"
        "3. Descargá el Excel final generado por la app."
    )

with st.container(border=True):
    st.subheader("Cargar archivos")
    archivos = st.file_uploader(
        "Arrastrá o seleccioná los 2 Excel",
        type=["xls", "xlsx", "xlsm", "xltx", "xltm"],
        accept_multiple_files=True,
        help="No hace falta renombrarlos. La app detecta cuál es la planilla y cuál es el fichaje.",
    )

    if archivos:
        st.write("**Archivos cargados:**")
        for archivo in archivos:
            st.write(f"- {archivo.name}")

    procesar = st.button("Procesar archivos", type="primary", use_container_width=True)

if procesar:
    try:
        if not archivos or len(archivos) < 2:
            st.error("Subí ambos archivos antes de procesar.")
        else:
            with st.spinner("Procesando planilla..."):
                resultado = procesar_desde_streamlit(archivos)

            st.success("Proceso completado.")
            with st.container(border=True):
                st.subheader("Archivos detectados")
                st.write(f"**Planilla base:** {resultado['archivo_planilla']}")
                st.write(f"**Fichajes:** {resultado['archivo_fichajes']}")
                st.dataframe(resultado["diagnostico"], use_container_width=True, hide_index=True)

            with st.container(border=True):
                st.subheader("Descargas")
                salidas = resultado["salidas"]
                if len(salidas) == 1:
                    nombre, contenido = next(iter(salidas.items()))
                    st.download_button(
                        label=f"Descargar {nombre}",
                        data=contenido,
                        file_name=nombre,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                else:
                    zip_bytes = io.BytesIO()
                    with zipfile.ZipFile(zip_bytes, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                        for nombre, contenido in salidas.items():
                            zf.writestr(nombre, contenido)
                    zip_bytes.seek(0)
                    st.download_button(
                        label="Descargar resultados (.zip)",
                        data=zip_bytes.getvalue(),
                        file_name="planillas_procesadas.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )

            with st.expander("Qué ajusta esta versión"):
                st.write(
                    "Reubica fórmulas diarias según día real, conserva la lógica de la plantilla, "
                    "calcula ausencias/enfermedad/ART/vacaciones con 8 h en semana y 4 h en sábado, "
                    "suma vianda por presencia y feriados según calendario argentino. Además, corrige el formato numérico de los resúmenes, marca nombres en escala de rojo según faltas y legajos en amarillo cuando el vínculo requiere revisión."
                )
    except Exception as e:
        st.error(f"No se pudo completar el proceso: {e}")
        st.exception(e)
