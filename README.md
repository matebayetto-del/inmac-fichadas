# INMAC · Planilla de fichajes v3

App Streamlit para procesar planillas base de Hemoterapia y Cubiertas.

## Archivos del repositorio
- `app.py`: interfaz web en Streamlit.
- `planilla_fichajes_inmac_v3.py`: motor principal.
- `requirements.txt`: dependencias.
- `logo_inmac.jpg`: logo de la empresa.
- `.gitignore`: exclusiones recomendadas.

## Qué hace
- Detecta automáticamente la planilla base y el archivo `Original` del fichador.
- Detecta el mes desde los fichajes.
- Reubica fórmulas diarias tomando las fórmulas originales de la plantilla base: laboral, sábado y domingo/feriado.
- Carga entrada/salida como horas reales de Excel.
- Marca `AUSENTE` cuando no hay fichaje en día laborable.
- Actualiza columnas de resumen: Ausencia, Enfermedad, Vianda, ART, Feriado y Vacaciones.
- Agrega columnas ART/FERIADO/VACACIONES si la plantilla no las trae.

## Publicación en Streamlit
1. Subir estos archivos al repositorio de GitHub.
2. Entrar a Streamlit Community Cloud.
3. Elegir el repo y publicar usando `app.py` como archivo principal.
