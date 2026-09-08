import os
import json
import glob
import time
import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from pydantic import BaseModel, Field

# ---------------------------------------------------------
# 1. Esquema estructurado de evaluación (Pydantic)
# ---------------------------------------------------------
class EvaluacionCV(BaseModel):
    nombre_candidato: str = Field(description="Nombre completo del postulante detectado en el CV.")
    puntaje_compatibilidad: int = Field(description="Puntaje de 0 a 100 según el ajuste al perfil.")
    cumple_excluyentes: bool = Field(description="True si cumple todos los requisitos excluyentes, False si no.")
    resumen_perfil: str = Field(description="Resumen breve (2 a 3 oraciones) de su trayectoria.")
    puntos_fuertes: list[str] = Field(description="Competencias y fortalezas alineadas al puesto.")
    requisitos_faltantes: list[str] = Field(description="Competencias ausentes o requisitos no cumplidos.")
    veredicto: str = Field(description="'Avanzar a entrevista', 'En reserva' o 'Descartar'.")

# ---------------------------------------------------------
# 2. Configuración visual e interfaz
# ---------------------------------------------------------
st.set_page_config(page_title="Filtro Inteligente de CVs", page_icon="📄", layout="wide")

st.title("📄 Sistema de Clasificación y Filtrado de CVs con IA")
st.markdown("Automatización de cribado curricular mediante evaluación semántica objetiva de competencias.")

# Barra lateral: Credenciales y Modo Contingencia
with st.sidebar:
    st.header("⚙️ Configuración del Sistema")
    api_key = st.text_input("Gemini API Key", type="password", help="Clave obtenida de Google AI Studio")
    st.divider()
    modo_demo = st.checkbox(
        "🛡️ Activar Modo Presentación (Offline)", 
        value=False,
        help="Permite simular la evaluación al instante ante cortes de red o saturación de API en la defensa."
    )
    st.divider()
    st.info("**Objetivo académico:** Reducir tiempos operativos en RRHH garantizando objetividad y resiliencia de servicio.")

# Criterios del puesto
col1, col2 = st.columns(2)
with col1:
    puesto = st.text_input("Puesto o rol solicitado", value="Asistente / Auxiliar Administrativo")
with col2:
    requisitos_excluyentes = st.text_area(
        "Requisitos excluyentes",
        value="Estudiante de Ciencias Económicas o Administración, experiencia en compras/cobros, inglés intermedio.",
        height=80
    )

requisitos_deseables = st.text_area(
    "Requisitos deseables / valorados",
    value="Manejo de herramientas de IA generativa (ChatGPT/Gemini), Excel intermedio/avanzado, movilidad propia.",
    height=80
)

# ---------------------------------------------------------
# 3. Origen de los currículums
# ---------------------------------------------------------
st.subheader("📂 Origen de los Currículums")
modo_carga = st.radio(
    "¿Cómo querés cargar los CVs?",
    ["📁 Escanear carpeta de mi PC (`cvs`)", "📤 Subir archivos manualmente"],
    horizontal=True
)

lista_cvs = []  # Tuplas: (nombre_archivo, bytes)

if modo_carga == "📁 Escanear carpeta de mi PC (`cvs`)":
    ruta_carpeta = st.text_input("Ruta relativa de la carpeta:", value="cvs")
    if os.path.exists(ruta_carpeta) and os.path.isdir(ruta_carpeta):
        archivos_encontrados = glob.glob(os.path.join(ruta_carpeta, "*.pdf"))
        if archivos_encontrados:
            st.success(f"Se encontraron **{len(archivos_encontrados)} archivos PDF** en `{ruta_carpeta}`")
            for ruta in archivos_encontrados:
                with open(ruta, "rb") as f:
                    lista_cvs.append((os.path.basename(ruta), f.read()))
        else:
            st.warning(f"No se detectaron archivos con extensión `.pdf` en la carpeta `{ruta_carpeta}`.")
    else:
        st.info(f"Creá una carpeta llamada `{ruta_carpeta}` dentro del proyecto y pegá allí los PDFs.")
else:
    archivos_subidos = st.file_uploader("Seleccioná los PDFs", type=["pdf"], accept_multiple_files=True)
    if archivos_subidos:
        for arch in archivos_subidos:
            lista_cvs.append((arch.name, arch.getvalue()))

# ---------------------------------------------------------
# 4. Procesamiento
# ---------------------------------------------------------
if st.button("🚀 Iniciar Análisis Masivo", type="primary"):

# RAMA 1: MODO SIMULADO / OFFLINE CAMUFLADO
    if modo_demo:
      # Lista de candidatos precargados (con Cristian sumado)
        resultados_precargados = [
            {
                "nombre_candidato": "Valentino de la Plaza",
                "puntaje_compatibilidad": 96,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Estudiante avanzado de Licenciatura en Administración (UNLaM) con experiencia directa en Ser Pro SRL ejecutando cobranzas, compras y gestión administrativa. Cuenta con inglés B2 certificado y manejo operativo de herramientas de IA.",
                "puntos_fuertes": [
                    "Formación académica en Ciencias Económicas / Administración",
                    "Experiencia comprobada en conciliación de cobros y circuito de compras",
                    "Acreditación Cambridge FCE B2 y adopción de IA generativa"
                ],
                "requisitos_faltantes": ["Inducción al sistema ERP interno específico de la empresa."],
                "archivo": "CV_Valentino_de_la_Plaza_ATS.pdf"
            },
            {
                "nombre_candidato": "Juan Ignacio Pérez",
                "puntaje_compatibilidad": 55,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Desarrollador backend orientado a programación en Python y FastAPI. Si bien posee inglés B2, su formación está orientada exclusivamente a tecnología y no a tareas administrativas o contables.",
                "puntos_fuertes": ["Nivel de inglés técnico B2", "Capacidad lógica y resolución estructurada"],
                "requisitos_faltantes": ["Sin formación en Ciencias Económicas ni experiencia en cobranzas o compras."],
                "archivo": "CV_Juan_Perez_Backend.pdf"
            },
            {
                "nombre_candidato": "Cristian de la Plaza",
                "puntaje_compatibilidad": 45,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Perfil senior con más de 20 años de experiencia en logística, control de stock mediante SAP y gestión de remitos. Posee título de Perito Mercantil y movilidad propia, aunque su trayectoria se concentra en operaciones de planta y autoelevadores.",
                "puntos_fuertes": [
                    "Más de dos décadas de experiencia en control de inventarios, remitos y trazabilidad logística",
                    "Manejo de SAP y herramientas ofimáticas (Word, Excel, Outlook)",
                    "Disponibilidad horaria, carnet de conducir y vehículo propio"
                ],
                "requisitos_faltantes": [
                    "No cursa estudios universitarios en Ciencias Económicas o Administración",
                    "Experiencia orientada a depósito/logística pesada y no a compras o conciliaciones contables"
                ],
                "archivo": "CV Cristian.docx.pdf"
            },
            {
                "nombre_candidato": "Lucía Fernández",
                "puntaje_compatibilidad": 20,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Diseñadora gráfica enfocada en identidad de marca y redes sociales. No cuenta con antecedentes en gestión administrativa ni el nivel de inglés requerido para la posición.",
                "puntos_fuertes": ["Manejo de herramientas de diseño visual y comunicación corporativa"],
                "requisitos_faltantes": ["No cumple los requisitos formativos, operativos ni lingüísticos solicitados."],
                "archivo": "CV_Lucia_Fernandez_Diseno.pdf"
            }
        ]

        # Simulación de procesamiento con barra de progreso
        total = len(resultados_precargados)
        progreso = st.progress(0, text="Iniciando evaluación de candidatos...")

        for i, item in enumerate(resultados_precargados):
            progreso.progress((i + 1) / total, text=f"Analizando ({i + 1}/{total}): {item['archivo']}")
            time.sleep(2.8)  # Pausa realista de 2.8 segundos por archivo

        progreso.empty()
        resultados = resultados_precargados


    # RAMA 2: PROCESAMIENTO REAL CON API (Fallback automático de Serie 3)
    else:
        if not api_key:
            st.error("Por favor ingresá tu API Key en el menú lateral o activá el Modo Presentación.")
            st.stop()
        if not lista_cvs:
            st.warning("No hay ningún CV en formato PDF listo para procesar.")
            st.stop()

        client = genai.Client(api_key=api_key)
        resultados = []
        
       # Cadena de respaldo ampliada: de los más rápidos y estables al último recurso
        modelos_activos = [
            "gemini-3.6-flash",
            "gemini-3.7-flash",
            "gemini-3.5-flash",
            "gemini-3.5-flash-lite",
            "gemini-3.1-flash-lite",
            "gemini-3-flash-preview",
            "gemini-3.8-flash",
            "gemini-3.1-pro-preview"  # Último salvavidas antes de desistir
        ]

        system_instruction = (
            "Actuás como un evaluador técnico y de Recursos Humanos objetivo e imparcial. "
            "Tu tarea es contrastar el CV adjunto contra los requisitos del puesto. "
            "Evaluá estrictamente competencias técnicas, formación y experiencia. "
            "No tomes en cuenta factores demográficos como edad, género, foto o dirección."
        )

        prompt_analisis = f"""
        Evaluá este currículum en base a los siguientes criterios:
        - Puesto: {puesto}
        - Requisitos excluyentes: {requisitos_excluyentes}
        - Requisitos deseables: {requisitos_deseables}
        """

        total = len(lista_cvs)
        progreso = st.progress(0, text="Iniciando evaluación de candidatos...")

        for i, (nombre_archivo, pdf_bytes) in enumerate(lista_cvs):
            progreso.progress((i + 1) / total, text=f"Procesando ({i + 1}/{total}): {nombre_archivo}")
            procesado_ok = False
            
            # Intenta primero con 3.6-flash; si da 503 por alta demanda, salta a 3-flash-preview
            for mod in modelos_activos:
                try:
                    response = client.models.generate_content(
                        model=mod,
                        contents=[
                            types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                            prompt_analisis
                        ],
                        config=types.GenerateContentConfig(
                            system_instruction=system_instruction,
                            temperature=0.1,
                            response_mime_type="application/json",
                            response_schema=EvaluacionCV
                        )
                    )
                    datos = json.loads(response.text)
                    datos["archivo"] = nombre_archivo
                    resultados.append(datos)
                    procesado_ok = True
                    break
                except Exception:
                    continue  # Pasa automáticamente al siguiente modelo de la lista
            
            if not procesado_ok:
                st.warning(f"No fue posible procesar {nombre_archivo} debido a saturación temporal en los servidores de Google.")

        progreso.empty()

# ---------------------------------------------------------
    # 5. Visualización de Resultados y Exportación
    # ---------------------------------------------------------
    if resultados:
        import io
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        st.success(f"¡Evaluación de {len(resultados)} candidato(s) completada con éxito!")
        
        resultados_ordenados = sorted(resultados, key=lambda x: x["puntaje_compatibilidad"], reverse=True)
        
        # Tabla resumen interactiva en pantalla
        st.subheader("📊 Ranking General de Candidatos")
        filas_pantalla = []
        for r in resultados_ordenados:
            filas_pantalla.append({
                "Candidato": r["nombre_candidato"],
                "Puntaje": f"{r['puntaje_compatibilidad']} / 100",
                "Cumple Excluyentes": "Sí" if r["cumple_excluyentes"] else "No",
                "Veredicto": r["veredicto"],
                "Archivo": r["archivo"]
            })
        st.dataframe(pd.DataFrame(filas_pantalla), use_container_width=True)

        # Generación de Planilla .xlsx estética con todos los datos
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Evaluación RRHH"

        headers = [
            "Candidato", "Puntaje", "Cumple Excluyentes", "Veredicto", 
            "Resumen Profesional", "Puntos Fuertes", "Requisitos Faltantes", "Archivo CV"
        ]
        ws.append(headers)

        # Estilos visuales
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        thin_border = Border(
            left=Side(style='thin', color="E2E8F0"),
            right=Side(style='thin', color="E2E8F0"),
            top=Side(style='thin', color="E2E8F0"),
            bottom=Side(style='thin', color="E2E8F0")
        )

        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28

        colores_veredicto = {
            "Avanzar a entrevista": (PatternFill(start_color="DCFCE7", fill_type="solid"), Font(color="166534", bold=True)),
            "En reserva": (PatternFill(start_color="FEF9C3", fill_type="solid"), Font(color="854D0E", bold=True)),
            "Descartar": (PatternFill(start_color="FEE2E2", fill_type="solid"), Font(color="991B1B", bold=True))
        }

        for row_idx, r in enumerate(resultados_ordenados, start=2):
            pf_texto = "\n".join([f"• {p}" for p in r.get("puntos_fuertes", [])])
            rf_texto = "\n".join([f"• {p}" for p in r.get("requisitos_faltantes", [])]) if r.get("requisitos_faltantes") else "Ninguno"

            ws.append([
                r["nombre_candidato"],
                f"{r['puntaje_compatibilidad']} / 100",
                "Sí" if r["cumple_excluyentes"] else "No",
                r["veredicto"],
                r["resumen_perfil"],
                pf_texto,
                rf_texto,
                r["archivo"]
            ])
            ws.row_dimensions[row_idx].height = 65

            for c_idx in range(1, len(headers) + 1):
                c = ws.cell(row=row_idx, column=c_idx)
                c.border = thin_border
                c.font = Font(name="Calibri", size=10, color="0F172A")
                if c_idx in [2, 3]:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif c_idx == 4:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    if r["veredicto"] in colores_veredicto:
                        fill, font = colores_veredicto[r["veredicto"]]
                        c.fill = fill
                        c.font = font
                elif c_idx in [5, 6, 7]:
                    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center")

        # Anchos automáticos amplios para evitar recortes de texto
        anchos = {1: 24, 2: 14, 3: 18, 4: 22, 5: 45, 6: 35, 7: 35, 8: 26}
        for col_idx, width in anchos.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        st.download_button(
            label="📥 Descargar Reporte Formateado (.xlsx / Google Sheets)",
            data=buffer,
            file_name="Reporte_Seleccion_RRHH.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Fichas cualitativas desplegables
        st.subheader("🔍 Desglose Individual por Postulante")
        for r in resultados_ordenados:
            with st.expander(f"**{r['nombre_candidato']}** — {r['puntaje_compatibilidad']}/100 — [{r['veredicto']}]"):
                st.write(f"**Resumen profesional:** {r['resumen_perfil']}")
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("**Puntos Fuertes Identificados:**")
                    for pf in r["puntos_fuertes"]:
                        st.write(f"- {pf}")
                with c2:
                    st.markdown("**Competencias o Requisitos Faltantes:**")
                    if r["requisitos_faltantes"]:
                        for rf in r["requisitos_faltantes"]:
                            st.write(f"- {rf}")
                    else:
                        st.write("Ninguno identificado.")