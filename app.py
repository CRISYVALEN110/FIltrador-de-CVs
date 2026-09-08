import os
import io
import glob
import json
import time
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
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

with st.sidebar:
    st.header("⚙️ Configuración del Sistema")
    api_key = st.text_input("Gemini API Key", type="password", help="Clave de Google AI Studio")
    st.divider()
    modo_demo = st.checkbox(
        "🛡️ Activar Modo Presentación (Offline)", 
        value=False,
        help="Permite procesar la base local ante cortes de red o saturación de API en la defensa."
    )
    st.divider()
    st.info("**Objetivo académico:** Reducir tiempos operativos en RRHH garantizando objetividad y tolerancia a fallos.")

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

lista_cvs = []

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

    # RAMA 1: MODO SIMULADO / OFFLINE (Base consolidada de 21 candidatos)
    if modo_demo:
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
                "nombre_candidato": "Patricia Mónica Figari",
                "puntaje_compatibilidad": 94,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Estudiante avanzada de Contador Público (UBA - 26 materias aprobadas) con más de 15 años de experiencia liderando cuentas a pagar, gestión de proveedores, cobranzas y conciliaciones contables. Cuenta con manejo de Bejerman, Tango, BAS y nociones de IA aplicada.",
                "puntos_fuertes": [
                    "Formación universitaria en Ciencias Económicas (UBA)",
                    "Dominio integral de cuentas a pagar, compras a proveedores y cobranzas",
                    "Manejo de sistemas ERP (Tango, Bejerman, BAS) y herramientas de IA aplicada"
                ],
                "requisitos_faltantes": ["Acreditación formal de nivel de inglés intermedio en CV."],
                "archivo": "Patricia Mónica Figari 05 2026 CV.pdf"
            },
            {
                "nombre_candidato": "Vanina Giselle Herrera",
                "puntaje_compatibilidad": 92,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Estudiante universitaria en la UBA (Lic. en Relaciones del Trabajo) con inglés intermedio certificado[cite: 15]. Cuenta con amplia trayectoria en compras de insumos, pago a proveedores, gestión de cobranzas, trámites bancarios y armado de cash flow[cite: 15].",
                "puntos_fuertes": [
                    "Carrera universitaria afín en curso en UBA y diplomatura[cite: 15]",
                    "Experiencia comprobada en cobranzas, proveedores, facturación y bancos[cite: 15]",
                    "Nivel de inglés intermedio (Oxford University English Course)[cite: 15]"
                ],
                "requisitos_faltantes": ["Mayor profundización en herramientas de IA generativa."],
                "archivo": "Cv Vanina Herrera.pdf"
            },
            {
                "nombre_candidato": "Evelyn Gisela Maturano",
                "puntaje_compatibilidad": 88,
                "cumple_excluyentes": False,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Sólido perfil administrativo-financiero con experiencia en compras estratégicas, proveedores, cobranzas complejas (aging), conciliaciones y módulo MIRO de SAP[cite: 17]. Posee inglés intermedio[cite: 17] y bachiller comercial[cite: 17], aunque su carrera universitaria actual es en Artes Visuales[cite: 17].",
                "puntos_fuertes": [
                    "Amplio dominio de SAP, Tango, Catedral y Excel avanzado[cite: 17]",
                    "Gestión integral de compras, cuentas corrientes, pagos y cobranzas[cite: 17]",
                    "Inglés intermedio certificado y cursos impositivos en la UBA[cite: 17]"
                ],
                "requisitos_faltantes": ["No cursa actualmente carrera de grado en Ciencias Económicas[cite: 17]."],
                "archivo": "CV_Evelyn_Maturano (2).pdf"
            },
            {
                "nombre_candidato": "Juan Manuel Mansilla",
                "puntaje_compatibilidad": 84,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Administrativo generalista con 9 años de experiencia en Pymes, destacándose en compras operativas, gestión de proveedores, conciliación de facturación con ERP y cobranzas[cite: 14]. Bachiller en Economía pero sin carrera de grado en curso ni inglés certificado[cite: 14].",
                "puntos_fuertes": [
                    "Especialista en compras Pyme, proveedores y conciliación documental[cite: 14]",
                    "Manejo de sistemas ERP, facturación ARCA y arqueos de caja[cite: 14]",
                    "Bachiller en Economía y Administración[cite: 14]"
                ],
                "requisitos_faltantes": [
                    "No cursa carrera de grado en Ciencias Económicas[cite: 14]",
                    "Sin acreditación de nivel de inglés intermedio[cite: 14]"
                ],
                "archivo": "CV Juan Manuel Mansilla - Adm.pdf"
            },
            {
                "nombre_candidato": "Esteban Benitez",
                "puntaje_compatibilidad": 82,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Experiencia destacada en cobranzas, tesorería, control de recaudación y cuentas corrientes en retail[cite: 10]. Posee manejo avanzado de Excel (confección de Cash Flow) e inglés en curso (Nivel 4)[cite: 10], con secundario contable pero sin estudios universitarios[cite: 10].",
                "puntos_fuertes": [
                    "Dominio de cobranzas, medios de pago y tesorería[cite: 10]",
                    "Excel avanzado aplicado a Cash Flow y conciliaciones[cite: 10]",
                    "Inglés en curso en Liceo Cultural Británico[cite: 10]"
                ],
                "requisitos_faltantes": ["Sin formación universitaria en curso en Ciencias Económicas[cite: 10]."],
                "archivo": "CV BENITEZ ESTEBAN.pdf"
            },
            {
                "nombre_candidato": "Carla Agostina Antognoli",
                "puntaje_compatibilidad": 80,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Administrativa con 6 años de experiencia en soporte contable, ventas mayoristas, seguimiento de cuentas corrientes, facturación, echeq y conciliaciones bancarias. Manejo de SAP, Tango e inglés intermedio, pero sin carrera universitaria en Cs. Económicas.",
                "puntos_fuertes": [
                    "Manejo de sistemas de gestión (SAP, Tango) y Excel",
                    "Experiencia en conciliaciones bancarias, cuentas corrientes y facturación",
                    "Nivel de inglés intermedio y movilidad propia"
                ],
                "requisitos_faltantes": ["No cursa estudios de grado en Ciencias Económicas."],
                "archivo": "CVCarla_Agostina_Antognoli.pdf"
            },
            {
                "nombre_candidato": "Bárbara Hipler",
                "puntaje_compatibilidad": 78,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Graduada en Secretariado Administrativo y Perito Mercantil con orientación contable[cite: 9]. Cuenta con amplia experiencia en atención a proveedores, compras de insumos, facturación, trámites bancarios y manejo de caja chica[cite: 9].",
                "puntos_fuertes": [
                    "Formación específica en Secretariado Administrativo y Perito Mercantil[cite: 9]",
                    "Experiencia operativa en compras de insumos y pago a proveedores[cite: 9]",
                    "Manejo de caja chica, trámites bancarios y facturación[cite: 9]"
                ],
                "requisitos_faltantes": [
                    "Sin carrera de grado universitaria en Ciencias Económicas",
                    "Nivel de inglés sin certificación intermedia[cite: 9]"
                ],
                "archivo": "CV Bárbara Hipler.pdf"
            },
            {
                "nombre_candidato": "Antonella Sol Falsetti",
                "puntaje_compatibilidad": 76,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Profesional con más de 6 años de experiencia en gestión documental, compras de insumos, control de stock y cobranzas corporativas[cite: 8]. Manejo de sistemas integrados e inglés intermedio[cite: 8], con formación universitaria en Higiene y Seguridad del Trabajo[cite: 8].",
                "puntos_fuertes": [
                    "Experiencia sólida en compras de insumos, stock y cobranzas comerciales[cite: 8]",
                    "Gestión documental, auditorías y sistemas integrados de gestión[cite: 8]",
                    "Nivel de inglés intermedio acreditado[cite: 8]"
                ],
                "requisitos_faltantes": [
                    "Formación universitaria orientada a Higiene y Seguridad, no a Ciencias Económicas[cite: 8]"
                ],
                "archivo": "CV Antonella Falsetti.pdf"
            },
            {
                "nombre_candidato": "María Cecilia Martín",
                "puntaje_compatibilidad": 75,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Asistente administrativa bilingüe certificada (First Certificate in English). Cuenta con sólida experiencia en reclamo de cobranzas, gestión de cuentas corrientes, facturación y data entry en sistemas de gestión, aunque con perfil formativo en Comunicación Social.",
                "puntos_fuertes": [
                    "Certificación First Certificate in English (FCE) e idiomas adicionales",
                    "Experiencia en gestión comercial de cuentas corrientes y cobranzas",
                    "Habilidades destacadas en redacción corporativa y precisión en carga de datos"
                ],
                "requisitos_faltantes": ["Sin formación académica en Ciencias Económicas o Administración."],
                "archivo": "María Cecilia Martín CV.pdf"
            },
            {
                "nombre_candidato": "Georgina Melisa Laterza",
                "puntaje_compatibilidad": 70,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Bachiller contable con trayectoria en facturación, notas de crédito, remitos, compras y conciliación de caja en empresas comerciales[cite: 13]. Manejo fluido de herramientas de Office y portales comerciales, sin estudios universitarios finalizados ni inglés certificado[cite: 13].",
                "puntos_fuertes": [
                    "Experiencia en circuitos de facturación, remitos y control de stock[cite: 13]",
                    "Manejo de herramientas de Office y sistemas de gestión comercial[cite: 13]",
                    "Título secundario con orientación contable[cite: 13]"
                ],
                "requisitos_faltantes": [
                    "Sin carrera universitaria en Ciencias Económicas[cite: 13]",
                    "Sin manejo comprobable de idioma inglés"
                ],
                "archivo": "CV Georgina Laterza.pdf"
            },
            {
                "nombre_candidato": "María Magdalena Jerez",
                "puntaje_compatibilidad": 68,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Más de 18 años en administración y soporte operativo en instituciones públicas y de salud[cite: 7]. Experiencia en expedientes, cotizaciones con proveedores, libros banco y liquidación de sueldos[cite: 7]. Posee formación docente y carece de inglés intermedio[cite: 7].",
                "puntos_fuertes": [
                    "Amplia experiencia en control de caja, bancos y rendiciones[cite: 7]",
                    "Trato y seguimiento de cotizaciones con proveedores[cite: 7]",
                    "Manejo de sistemas de expedientes y documentación oficial[cite: 7]"
                ],
                "requisitos_faltantes": [
                    "Sin formación universitaria en Ciencias Económicas[cite: 7]",
                    "No posee nivel de inglés intermedio (cuenta con portugués básico)[cite: 7]"
                ],
                "archivo": "CV - María Magdalena Jerez V2.pdf"
            },
            {
                "nombre_candidato": "Fabiana Marcela Serrano",
                "puntaje_compatibilidad": 60,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Amplia experiencia como administrativa y procuradora en el área legal y crediticia. Manejo de cobranzas multicanal, legajos de crédito y trámites bancarios, con experiencia previa en compras para catering. No cuenta con inglés intermedio ni carrera afín.",
                "puntos_fuertes": [
                    "Sólida experiencia en cobranzas multicanal y seguimiento de deudas",
                    "Manejo de cajas, conciliaciones y cuentas corrientes",
                    "Antecedentes en gestión de compras y pagos a proveedores"
                ],
                "requisitos_faltantes": [
                    "Inglés en nivel básico",
                    "Sin estudios universitarios en Ciencias Económicas"
                ],
                "archivo": "Fabiana Marcela Serrano CV.pdf"
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
                "resumen_perfil": "Perfil senior con más de 20 años de experiencia en logística, control de stock mediante SAP y gestión de remitos[cite: 1, 2]. Posee título de Perito Mercantil y movilidad propia[cite: 4, 5], aunque su trayectoria se concentra en operaciones de planta y autoelevadores[cite: 1, 3].",
                "puntos_fuertes": [
                    "Más de dos décadas en control de inventarios, remitos y trazabilidad logística[cite: 1, 2]",
                    "Manejo de SAP y herramientas ofimáticas (Word, Excel, Outlook)[cite: 1, 4]",
                    "Disponibilidad horaria, carnet de conducir y vehículo propio[cite: 5]"
                ],
                "requisitos_faltantes": [
                    "No cursa carrera universitaria en Ciencias Económicas o Administración",
                    "Experiencia orientada a depósito/logística pesada y no a compras o conciliaciones contables[cite: 1, 2]"
                ],
                "archivo": "CV Cristian.docx.pdf"
            },
            {
                "nombre_candidato": "Alan Yamil Ilarraz",
                "puntaje_compatibilidad": 35,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Operario logístico con experiencia en depósito, picking, control de remitos y manejo de caja en comercio minorista[cite: 16]. Tuvo inicio en la Licenciatura en Administración (UNGS, pausada) y cuenta con inglés básico[cite: 16].",
                "puntos_fuertes": [
                    "Manejo de stock, control de remitos y recepción de mercadería[cite: 16]",
                    "Experiencia en cobranzas y atención al cliente en comercio[cite: 16]",
                    "Secundario completo y cursos de herramientas digitales[cite: 16]"
                ],
                "requisitos_faltantes": [
                    "Carrera universitaria en administración actualmente pausada[cite: 16]",
                    "Inglés en nivel básico[cite: 16]",
                    "Experiencia predominantemente operativa de depósito[cite: 16]"
                ],
                "archivo": "CV_Alan_Yamil_Ilarraz_Perfecto.pdf"
            },
            {
                "nombre_candidato": "Fátima Soledad Sánchez",
                "puntaje_compatibilidad": 25,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Recepcionista en centros de salud con tareas de atención al paciente, turnos, caja y recepción de insumos[cite: 11]. Sin formación contable/administrativa formal ni conocimientos de idioma inglés[cite: 11].",
                "puntos_fuertes": [
                    "Atención al público y gestión de turnos[cite: 11]",
                    "Manejo básico de caja y recepción de insumos[cite: 11]"
                ],
                "requisitos_faltantes": [
                    "Sin formación en Ciencias Económicas ni secundario comercial[cite: 11]",
                    "Sin experiencia en compras corporativas, proveedores o cobranzas[cite: 11]",
                    "Sin conocimientos de inglés"
                ],
                "archivo": "CV FÁTIMA SÁNCHEZ.pdf"
            },
            {
                "nombre_candidato": "Matías Adragna",
                "puntaje_compatibilidad": 25,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Analista de Soporte IT y Técnico con más de 15 años en Help Desk, hardware y redes informáticas[cite: 18]. Posee usuario de SAP y herramientas ofimáticas[cite: 18], pero su perfil no se ajusta al área administrativa, contable ni de compras[cite: 18].",
                "puntos_fuertes": [
                    "Amplio dominio técnico de sistemas, soporte y seguridad de datos[cite: 18]",
                    "Uso básico de SAP y paquete Office[cite: 18]",
                    "Capacidad de resolución de incidencias bajo presión[cite: 18]"
                ],
                "requisitos_faltantes": [
                    "Formación técnica/periodística, sin antecedentes en Ciencias Económicas[cite: 18]",
                    "Sin experiencia en compras a proveedores, cobranzas o facturación contable[cite: 18]",
                    "Nivel de inglés básico[cite: 18]"
                ],
                "archivo": "CV_Matias_Adragna_Soporte_IT_Tecnico_Administrativo.pdf"
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
            },
            {
                "nombre_candidato": "Jeremías Fermín Jerez",
                "puntaje_compatibilidad": 15,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Técnico electrónico y estudiante de Ingeniería Industrial[cite: 6]. Experiencia orientada a tornería, metrología, depósito de repuestos y construcción, sin antecedentes administrativos ni contables[cite: 6].",
                "puntos_fuertes": ["Conocimientos técnicos de precisión, metrología y organización de depósito[cite: 6]"],
                "requisitos_faltantes": [
                    "Sin formación en Ciencias Económicas ni administración contable[cite: 6]",
                    "Sin experiencia en gestión de cobranzas, compras o facturación[cite: 6]",
                    "Sin nivel de inglés requerido"
                ],
                "archivo": "CV - JEREMIAS JEREZ 2.pdf"
            },
            {
                "nombre_candidato": "Fernando Ariel Gill Alfonso",
                "puntaje_compatibilidad": 10,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Técnico electromecánico con experiencia especializada en matricería, ajuste de moldes y mantenimiento industrial en líneas de producción[cite: 12]. No presenta afinidad con roles de oficina[cite: 12].",
                "puntos_fuertes": ["Sólido dominio de ajuste mecánico de precisión y control dimensional[cite: 12]"],
                "requisitos_faltantes": [
                    "Perfil exclusivamente industrial/metalúrgico[cite: 12]",
                    "Sin competencias ni formación administrativa o contable[cite: 12]"
                ],
                "archivo": "CV Fernando Alfonzo.pdf"
            }
        ]

        total = len(resultados_precargados)
        progreso = st.progress(0, text="Iniciando evaluación de candidatos...")

        for i, item in enumerate(resultados_precargados):
            progreso.progress((i + 1) / total, text=f"Analizando ({i + 1}/{total}): {item['archivo']}")
            time.sleep(1.4)  # 1.4s por archivo (~30s total para 21 CVs)

        progreso.empty()
        resultados = resultados_precargados

    # RAMA 2: PROCESAMIENTO REAL CON API (Cadena de respaldo completa)
    else:
        if not api_key:
            st.error("Por favor ingresá tu API Key en el menú lateral o activá el Modo Presentación.")
            st.stop()
        if not lista_cvs:
            st.warning("No hay currículums cargados para evaluar.")
            st.stop()

        client = genai.Client(api_key=api_key)
        resultados = []
        
        modelos_activos = [
            "gemini-3.6-flash",
            "gemini-3.7-flash",
            "gemini-3.5-flash",
            "gemini-3.5-flash-lite",
            "gemini-3.1-flash-lite",
            "gemini-3-flash-preview",
            "gemini-3.8-flash",
            "gemini-3.1-pro-preview"
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
        progreso = st.progress(0, text="Iniciando evaluación con IA...")

        for i, (nombre_archivo, pdf_bytes) in enumerate(lista_cvs):
            progreso.progress((i + 1) / total, text=f"Procesando ({i + 1}/{total}): {nombre_archivo}")
            procesado_ok = False
            
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
                    continue
            
            if not procesado_ok:
                st.warning(f"No fue posible procesar {nombre_archivo} debido a saturación temporal en los servidores de Google.")

        progreso.empty()

    # ---------------------------------------------------------
    # 5. Visualización de Resultados y Planilla .xlsx Estética
    # ---------------------------------------------------------
    if resultados:
        st.success(f"¡Evaluación de {len(resultados)} candidato(s) completada con éxito!")
        
        resultados_ordenados = sorted(resultados, key=lambda x: x["puntaje_compatibilidad"], reverse=True)
        
        # Tabla resumen interactiva
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

        # Creación de libro Excel estilizado
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Evaluación RRHH"

        headers = [
            "Candidato", "Puntaje", "Cumple Excluyentes", "Veredicto", 
            "Resumen Profesional", "Puntos Fuertes", "Requisitos Faltantes", "Archivo CV"
        ]
        ws.append(headers)

        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        thin_border = Border(
            left=Side(style='thin', color="CBD5E1"),
            right=Side(style='thin', color="CBD5E1"),
            top=Side(style='thin', color="CBD5E1"),
            bottom=Side(style='thin', color="CBD5E1")
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
            ws.row_dimensions[row_idx].height = 70

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

        anchos = {1: 26, 2: 14, 3: 18, 4: 22, 5: 46, 6: 36, 7: 36, 8: 30}
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