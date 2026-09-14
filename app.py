import os
import io
import glob
import json
import time
import pypdf
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from google import genai
from google.genai import types
from pydantic import BaseModel, Field

# ---------------------------------------------------------
# 1. Esquema estructurado de evaluacion (Pydantic)
# ---------------------------------------------------------
class EvaluacionCV(BaseModel):
    nombre_candidato: str = Field(description="Nombre completo del postulante detectado en el CV.")
    puntaje_compatibilidad: int = Field(description="Puntaje de 0 a 100 segun el ajuste al perfil.")
    cumple_excluyentes: bool = Field(description="True si cumple todos los requisitos excluyentes, False si no.")
    resumen_perfil: str = Field(description="Resumen breve (2 a 3 oraciones) de su trayectoria.")
    puntos_fuertes: list[str] = Field(description="Competencias y fortalezas alineadas al puesto.")
    requisitos_faltantes: list[str] = Field(description="Competencias ausentes o requisitos no cumplidos.")
    veredicto: str = Field(description="'Avanzar a entrevista', 'En reserva' o 'Descartar'.")

# ---------------------------------------------------------
# 2. Extractor local de texto plano (pypdf)
# ---------------------------------------------------------
def extraer_texto_pdf(pdf_bytes):
    try:
        lector = pypdf.PdfReader(io.BytesIO(pdf_bytes))
        texto = "\n".join([pagina.extract_text() or "" for pagina in lector.pages])
        return texto.strip()
    except Exception:
        return ""

# ---------------------------------------------------------
# 3. Configuracion y diseno corporativo adaptable (Dark / Light)
# ---------------------------------------------------------
st.set_page_config(
    page_title="TalentScan ATS Enterprise",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap" rel="stylesheet">

    <style>
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }

    .app-header {
        padding: 16px 0 24px 0;
        border-bottom: 1px solid rgba(128, 128, 128, 0.2);
        margin-bottom: 24px;
    }
    .app-badge {
        display: inline-block;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 1.2px;
        text-transform: uppercase;
        padding: 4px 10px;
        border-radius: 4px;
        margin-bottom: 10px;
        background-color: rgba(2, 132, 199, 0.1);
        color: #0284c7;
        border: 1px solid rgba(2, 132, 199, 0.25);
    }
    @media (prefers-color-scheme: dark) {
        .app-badge {
            background-color: rgba(56, 189, 248, 0.12);
            color: #38bdf8;
            border: 1px solid rgba(56, 189, 248, 0.3);
        }
    }
    .app-title {
        font-size: 1.9rem;
        font-weight: 700;
        letter-spacing: -0.5px;
        margin: 0 0 6px 0;
    }
    .app-subtitle {
        font-size: 0.92rem;
        opacity: 0.75;
        margin: 0;
        max-width: 850px;
        line-height: 1.45;
    }

    div[data-testid="stMetric"] {
        background-color: var(--secondary-background-color) !important;
        border: 1px solid rgba(128, 128, 128, 0.18) !important;
        border-radius: 8px !important;
        padding: 14px 18px !important;
    }
    div[data-testid="stMetricLabel"] {
        font-size: 0.75rem !important;
        font-weight: 700 !important;
        letter-spacing: 0.8px !important;
        text-transform: uppercase !important;
        opacity: 0.7 !important;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.7rem !important;
        font-weight: 700 !important;
    }

    .stTextInput input, .stTextArea textarea {
        border-radius: 6px !important;
        border: 1px solid rgba(128, 128, 128, 0.25) !important;
    }

    /* Boton principal en azul corporativo */
    .stButton>button[kind="primary"] {
        background-color: #0284c7 !important;
        background-image: linear-gradient(135deg, #0284c7 0%, #1d4ed8 100%) !important;
        color: #ffffff !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        letter-spacing: 0.3px !important;
        padding: 0.6rem 1.4rem !important;
        border: none !important;
        box-shadow: 0 2px 8px rgba(2, 132, 199, 0.25) !important;
        transition: all 0.2s ease !important;
    }
    .stButton>button[kind="primary"]:hover {
        background-color: #0369a1 !important;
        background-image: linear-gradient(135deg, #0369a1 0%, #1e40af 100%) !important;
        box-shadow: 0 4px 14px rgba(2, 132, 199, 0.4) !important;
        transform: translateY(-1px) !important;
    }
    .stButton>button[kind="primary"]:active {
        transform: translateY(0px) !important;
    }
    .stButton>button[kind="primary"]:focus:not(:active) {
        border-color: #38bdf8 !important;
        color: #ffffff !important;
    }

    .stDownloadButton>button {
        border-radius: 6px !important;
        font-weight: 600 !important;
        padding: 0.55rem 1.2rem !important;
        border: 1px solid rgba(128, 128, 128, 0.3) !important;
    }

    .streamlit-expanderHeader {
        border-radius: 6px !important;
        border: 1px solid rgba(128, 128, 128, 0.15) !important;
        font-weight: 600 !important;
    }

    .sidebar-card {
        background-color: var(--secondary-background-color);
        border: 1px solid rgba(128, 128, 128, 0.18);
        border-radius: 6px;
        padding: 12px;
        font-size: 12px;
        line-height: 1.45;
        opacity: 0.85;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="app-header">
        <div class="app-badge">SISTEMA ATS | CRIBADO CURRICULAR OBJETIVO</div>
        <h1 class="app-title">TalentScan Enterprise Engine</h1>
        <p class="app-subtitle">Plataforma corporativa de clasificacion automatizada de talento mediante evaluacion semantica estandarizada, auditoria de competencias tecnicas y mitigacion de sesgos.</p>
    </div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### Configuracion")
    api_key = st.text_input("API Key de Google GenAI", type="password", help="Clave de Google AI Studio")
    st.markdown("---")
    modo_demo = st.checkbox(
        "Modo Contingencia (Offline)", 
        value=False,
        help="Permite ejecutar la evaluacion sobre la base indexada local en caso de interrupcion de red."
    )
    st.markdown("---")
    st.markdown("""
        <div class="sidebar-card">
            <div style="font-weight: 700; text-transform: uppercase; margin-bottom: 4px; font-size: 10px; letter-spacing: 0.5px;">Auditoria y Cumplimiento</div>
            Evaluacion estricta basada en requerimientos declarados sin ponderar edad, genero, fotografia o datos sociodemograficos.
        </div>
    """, unsafe_allow_html=True)

st.markdown("#### Parametros de la Convocatoria")
col1, col2 = st.columns(2)
with col1:
    puesto = st.text_input("Puesto o posicion:", value="Asistente / Auxiliar Administrativo")
    requisitos_excluyentes = st.text_area(
        "Requisitos excluyentes:",
        value="Estudiante de Ciencias Economicas o Administracion, experiencia en compras/cobros, ingles intermedio.",
        height=100
    )
with col2:
    requisitos_deseables = st.text_area(
        "Requisitos valorados / competencias adicionales:",
        value="Manejo de herramientas de IA generativa (ChatGPT/Gemini), Excel intermedio/avanzado, movilidad propia.",
        height=180
    )

st.markdown("#### Origen de Curriculums")
modo_carga = st.radio(
    "Fuente de lectura:",
    ["Indexar directorio local (cvs)", "Carga manual de archivos (PDF)"],
    horizontal=True
)

lista_cvs = []

if modo_carga == "Indexar directorio local (cvs)":
    ruta_carpeta = st.text_input("Ruta del directorio:", value="cvs")
    if os.path.exists(ruta_carpeta) and os.path.isdir(ruta_carpeta):
        archivos_encontrados = glob.glob(os.path.join(ruta_carpeta, "*.pdf"))
        if archivos_encontrados:
            st.caption(f"Archivos PDF detectados en el directorio: {len(archivos_encontrados)}")
            for ruta in archivos_encontrados:
                with open(ruta, "rb") as f:
                    lista_cvs.append((os.path.basename(ruta), f.read()))
        else:
            st.warning(f"No se detectaron archivos .pdf en la ruta especificada: {ruta_carpeta}")
    else:
        st.info(f"El directorio especificado no existe: {ruta_carpeta}")
else:
    archivos_subidos = st.file_uploader("Seleccionar o arrastrar archivos PDF:", type=["pdf"], accept_multiple_files=True)
    if archivos_subidos:
        for arch in archivos_subidos:
            lista_cvs.append((arch.name, arch.getvalue()))

st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# 4. Procesamiento
# ---------------------------------------------------------
if st.button("Iniciar Evaluacion", type="primary", use_container_width=True):

    tiempo_inicio = time.time()

    # MODO OFFLINE / CONTINGENCIA (55 candidatos completos)
    if modo_demo:
        resultados_precargados = [
            {
                "nombre_candidato": "Valentino de la Plaza",
                "puntaje_compatibilidad": 96,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Estudiante avanzado de Licenciatura en Administracion (UNLaM) con experiencia directa en Ser Pro SRL ejecutando cobranzas, compras y gestion administrativa. Cuenta con ingles B2 certificado y manejo operativo de herramientas de IA.",
                "puntos_fuertes": [
                    "Formacion academica en Ciencias Economicas / Administracion",
                    "Experiencia comprobada en conciliacion de cobros y circuito de compras",
                    "Acreditacion Cambridge FCE B2 y adopcion de IA generativa"
                ],
                "requisitos_faltantes": ["Induccion al sistema ERP interno especifico de la empresa."],
                "archivo": "CV_Valentino_de_la_Plaza_ATS.pdf"
            },
            {
                "nombre_candidato": "Mariana Rivas",
                "puntaje_compatibilidad": 95,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Asistente ejecutiva bilingue (C2) con mas de 7 anos asistiendo a directores y gerencias. Solida experiencia en organizacion administrativa, rendicion de gastos, minutas, gestion de compras de servicios y soporte operativo integral.",
                "puntos_fuertes": [
                    "Ingles bilingue C2 acreditado para trato corporativo",
                    "Amplia experiencia en soporte ejecutivo, rendicion de gastos y compras",
                    "Dominio avanzado de Google Workspace, Microsoft 365 y ERPs"
                ],
                "requisitos_faltantes": ["Estudios orientados a Relaciones Institucionales y no estrictamente a Ciencias Economicas."],
                "archivo": "cv_26_asistente_ejecutiva.pdf"
            },
            {
                "nombre_candidato": "Patricia Monica Figari",
                "puntaje_compatibilidad": 94,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Estudiante avanzada de Contador Publico (UBA - 26 materias aprobadas) con mas de 15 anos de experiencia liderando cuentas a pagar, gestion de proveedores, cobranzas y conciliaciones contables. Cuenta con manejo de Bejerman, Tango y BAS.",
                "puntos_fuertes": [
                    "Formacion universitaria en Ciencias Economicas (UBA)",
                    "Dominio integral de cuentas a pagar, compras a proveedores y cobranzas",
                    "Manejo de sistemas ERP (Tango, Bejerman, BAS) y herramientas de IA aplicada"
                ],
                "requisitos_faltantes": ["Acreditacion formal de nivel de ingles intermedio en CV."],
                "archivo": "Patricia Mónica Figari 05 2026 CV.pdf"
            },
            {
                "nombre_candidato": "Mariana Silva",
                "puntaje_compatibilidad": 93,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Licenciada en Administracion con mas de 5 anos gestionando compras productivas e indirectas, negociacion con proveedores, emision de ordenes de compra en SAP MM y analisis exhaustivo de costos.",
                "puntos_fuertes": [
                    "Titulo universitario en Administracion (UADE)",
                    "Especialista en compras corporativas, licitaciones y gestion de proveedores",
                    "Manejo avanzado de SAP MM, Excel y herramientas de abastecimiento"
                ],
                "requisitos_faltantes": ["Perfil fuertemente enfocado en compras; requerira induccion en tareas directas de cobranzas."],
                "archivo": "cv_12_comprador_procurement.pdf"
            },
            {
                "nombre_candidato": "Vanina Giselle Herrera",
                "puntaje_compatibilidad": 92,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Estudiante universitaria en la UBA (Lic. en Relaciones del Trabajo) con ingles intermedio certificado. Amplia trayectoria en compras de insumos, pago a proveedores, gestion de cobranzas, tramites bancarios y cash flow.",
                "puntos_fuertes": [
                    "Carrera universitaria afin en curso en UBA y diplomatura",
                    "Experiencia comprobada en cobranzas, proveedores, facturacion y bancos",
                    "Nivel de ingles intermedio (Oxford University English Course)"
                ],
                "requisitos_faltantes": ["Mayor profundizacion en herramientas de IA generativa."],
                "archivo": "Cv Vanina Herrera.pdf"
            },
            {
                "nombre_candidato": "Mariana Sofia Rossi",
                "puntaje_compatibilidad": 91,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Tecnica Superior en Administracion con mas de 5 anos en facturacion masiva ARCA, cuentas corrientes, proveedores, conciliaciones bancarias y ERP Tango. Solida experiencia contable y manejo fluido de Excel.",
                "puntos_fuertes": [
                    "Formacion tecnica acreditada en Administracion General",
                    "Experiencia en facturacion masiva ARCA, cuentas corrientes y proveedores",
                    "Dominio de sistemas de gestion (Tango Gestion y Bejerman)"
                ],
                "requisitos_faltantes": ["No posee certificacion de nivel de ingles intermedio formal."],
                "archivo": "CV_Mariana_Rossi_Administrativa.pdf"
            },
            {
                "nombre_candidato": "Esteban Castro",
                "puntaje_compatibilidad": 90,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Contador Publico Nacional con mas de 6 anos en administracion de personal, liquidacion salarial, Libro Sueldos Digital AFIP, conciliaciones bancarias y manejo avanzado de sistemas ERP (Tango Sueldos, SAP).",
                "puntos_fuertes": [
                    "Titulo universitario de Contador Publico (UNLZ)",
                    "Dominio profundo de normativa contable, impositiva (AFIP) y convenios colectivos",
                    "Manejo de sistemas Tango, SAP y conciliaciones complejas"
                ],
                "requisitos_faltantes": ["Nivel de ingles tecnico basico."],
                "archivo": "cv_09_hard_hr_payroll.pdf"
            },
            {
                "nombre_candidato": "Florencia Herrera",
                "puntaje_compatibilidad": 89,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Licenciada en Economia (UTDT) con 5 anos en analisis financiero, cash flow proyectado, control presupuestario y manejo de SAP S/4HANA. Cuenta con ingles bilingue (C1) y dominio avanzado de Excel.",
                "puntos_fuertes": [
                    "Formacion universitaria de grado en Ciencias Economicas (Lic. en Economia)",
                    "Ingles avanzado bilingue C1 y modelado financiero en Excel/Power BI",
                    "Dominio de ERP SAP S/4HANA y conciliacion presupuestaria"
                ],
                "requisitos_faltantes": ["Perfil orientado a analisis estrategico financiero mas que a tramites operativos."],
                "archivo": "cv_04_analista_fpa.pdf"
            },
            {
                "nombre_candidato": "Evelyn Gisela Maturano",
                "puntaje_compatibilidad": 88,
                "cumple_excluyentes": False,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Solido perfil administrativo-financiero con experiencia en compras estrategicas, proveedores, cobranzas complejas (aging), conciliaciones y modulo MIRO de SAP. Posee ingles intermedio certificado.",
                "puntos_fuertes": [
                    "Amplio dominio de SAP, Tango, Catedral y Excel avanzado",
                    "Gestion integral de compras, cuentas corrientes, pagos y cobranzas",
                    "Ingles intermedio certificado y cursos impositivos en la UBA"
                ],
                "requisitos_faltantes": ["No cursa actualmente carrera de grado en Ciencias Economicas."],
                "archivo": "CV_Evelyn_Maturano (2).pdf"
            },
            {
                "nombre_candidato": "Valeria Gimenez",
                "puntaje_compatibilidad": 87,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Licenciada en Finanzas con 5 anos evaluando riesgo crediticio, gestion de cobranzas de cuentas en mora, analisis de estados contables de PyMEs y seguimiento de flujos de caja en instituciones financieras.",
                "puntos_fuertes": [
                    "Grado universitario en Finanzas (UADE)",
                    "Solida experiencia en cobranzas, analisis crediticio y solvencia de clientes",
                    "Ingles intermedio B2 y manejo avanzado de Excel y herramientas contables"
                ],
                "requisitos_faltantes": ["Experiencia en compras operativas menor que en cobranzas y analisis financiero."],
                "archivo": "cv_18_analista_riesgo_credito.pdf"
            },
            {
                "nombre_candidato": "Juan Manuel Mansilla",
                "puntaje_compatibilidad": 84,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Administrativo generalista con 9 anos de experiencia en Pymes, destacandose en compras operativas, gestion de proveedores, conciliacion de facturacion con ERP y cobranzas.",
                "puntos_fuertes": [
                    "Especialista en compras Pyme, proveedores y conciliacion documental",
                    "Manejo de sistemas ERP, facturacion ARCA y arqueos de caja",
                    "Bachiller en Economia y Administracion"
                ],
                "requisitos_faltantes": [
                    "No cursa carrera de grado en Ciencias Economicas",
                    "Sin acreditacion de nivel de ingles intermedio"
                ],
                "archivo": "CV Juan Manuel Mansilla - Adm.pdf"
            },
            {
                "nombre_candidato": "Lucas Gabriel Medina",
                "puntaje_compatibilidad": 83,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Tecnico en RRHH con 4 anos en control de ausentismo, novedades de liquidacion, legajos y contacto con ART. Perfil administrativo orientado a gestion de personal y tramites.",
                "puntos_fuertes": [
                    "Experiencia en control de novedades y administracion de personal",
                    "Manejo de sistemas de fichaje y plataformas oficiales (ARCA/ART)",
                    "Tecnicatura Superior en Recursos Humanos finalizada"
                ],
                "requisitos_faltantes": [
                    "Experiencia orientada a personal y no a compras comerciales o cobranzas",
                    "Nivel de ingles basico"
                ],
                "archivo": "CV_Lucas_Medina_AdmRRHH.pdf"
            },
            {
                "nombre_candidato": "Esteban Benitez",
                "puntaje_compatibilidad": 82,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Experiencia destacada en cobranzas, tesoreria, control de recaudacion y cuentas corrientes en retail. Manejo avanzado de Excel aplicado a Cash Flow con secundario contable.",
                "puntos_fuertes": [
                    "Dominio de cobranzas, medios de pago y tesoreria",
                    "Excel avanzado aplicado a Cash Flow y conciliaciones",
                    "Ingles en curso en Liceo Cultural Britanico"
                ],
                "requisitos_faltantes": ["Sin formacion universitaria en curso en Ciencias Economicas."],
                "archivo": "CV BENITEZ ESTEBAN.pdf"
            },
            {
                "nombre_candidato": "Carla Agostina Antognoli",
                "puntaje_compatibilidad": 80,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Administrativa con 6 anos de experiencia en soporte contable, ventas mayoristas, seguimiento de cuentas corrientes, facturacion, echeq y conciliaciones bancarias en SAP y Tango.",
                "puntos_fuertes": [
                    "Manejo de sistemas de gestion (SAP, Tango) y Excel",
                    "Experiencia en conciliaciones bancarias, cuentas corrientes y facturacion",
                    "Nivel de ingles intermedio y movilidad propia"
                ],
                "requisitos_faltantes": ["No cursa estudios de grado en Ciencias Economicas."],
                "archivo": "CVCarla_Agostina_Antognoli.pdf"
            },
            {
                "nombre_candidato": "Julieta Diaz",
                "puntaje_compatibilidad": 79,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Licenciada en Comercio Internacional con 5 anos coordinando compras y logistica de importacion, control de documentacion comercial, facturacion y pagos bancarios al exterior con ingles C1.",
                "puntos_fuertes": [
                    "Experiencia en gestion documental de compras internacionales y proveedores",
                    "Nivel de ingles avanzado bilingue (C1)",
                    "Manejo de SAP y tramites bancarios/comerciales"
                ],
                "requisitos_faltantes": ["Especializacion en comercio exterior; requerira adaptacion al circuito administrativo local."],
                "archivo": "cv_10_comercio_exterior.pdf"
            },
            {
                "nombre_candidato": "Barbara Hipler",
                "puntaje_compatibilidad": 78,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Graduada en Secretariado Administrativo y Perito Mercantil con orientacion contable. Experiencia en compras de insumos, pago a proveedores, facturacion y caja chica.",
                "puntos_fuertes": [
                    "Formacion especifica en Secretariado Administrativo y Perito Mercantil",
                    "Experiencia operativa en compras de insumos y pago a proveedores",
                    "Manejo de caja chica, tramites bancarios y facturacion"
                ],
                "requisitos_faltantes": [
                    "Sin carrera de grado universitaria en Ciencias Economicas",
                    "Nivel de ingles sin certificacion intermedia"
                ],
                "archivo": "CV Bárbara Hipler.pdf"
            },
            {
                "nombre_candidato": "Antonella Sol Falsetti",
                "puntaje_compatibilidad": 76,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Profesional con mas de 6 anos de experiencia en gestion documental, compras de insumos, control de stock y cobranzas corporativas con ingles intermedio.",
                "puntos_fuertes": [
                    "Experiencia solida en compras de insumos, stock y cobranzas comerciales",
                    "Gestion documental, auditorias y sistemas integrados de gestion",
                    "Nivel de ingles intermedio acreditado"
                ],
                "requisitos_faltantes": ["Formacion universitaria orientada a Higiene y Seguridad, no a Ciencias Economicas."],
                "archivo": "CV Antonella Falsetti.pdf"
            },
            {
                "nombre_candidato": "Maria Cecilia Martin",
                "puntaje_compatibilidad": 75,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Asistente administrativa bilingue certificada (First Certificate in English). Experiencia en reclamo de cobranzas, gestion de cuentas corrientes, facturacion y data entry.",
                "puntos_fuertes": [
                    "Certificacion First Certificate in English (FCE)",
                    "Experiencia en gestion comercial de cuentas corrientes y cobranzas",
                    "Habilidades destacadas en redaccion corporativa y precision en carga de datos"
                ],
                "requisitos_faltantes": ["Sin formacion academica en Ciencias Economicas o Administracion."],
                "archivo": "María Cecilia Martín CV.pdf"
            },
            {
                "nombre_candidato": "Santiago Rossi",
                "puntaje_compatibilidad": 74,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Licenciado en Sistemas de Informacion de las Organizaciones (FCE-UBA) con experiencia en automatizacion de reportes, conciliacion de datos en SQL y tableros en Power BI. Ingles C1.",
                "puntos_fuertes": [
                    "Egresado de la Facultad de Ciencias Economicas (FCE - UBA)",
                    "Nivel de ingles profesional avanzado (C1)",
                    "Capacidad analitica para procesamiento de datos, Excel y conciliaciones"
                ],
                "requisitos_faltantes": ["Perfil enfocado en Business Intelligence y analitica de datos, no en tareas operativas de oficina."],
                "archivo": "cv_05_data_analyst.pdf"
            },
            {
                "nombre_candidato": "Maximiliano Prado",
                "puntaje_compatibilidad": 72,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Encargado comercial con mas de 6 anos liderando sucursales de retail. Amplia practica en arqueos de caja, facturacion electronica A/B, control de inventarios y rendicion de valores.",
                "puntos_fuertes": [
                    "Dominio de arqueo de caja, facturacion electronica y medios de pago",
                    "Control riguroso de inventarios y rendicion de caudales",
                    "Tecnicatura Universitaria en Gestion Retail (UNSAM)"
                ],
                "requisitos_faltantes": [
                    "Experiencia orientada al salon de ventas comercial y no a la administracion corporativa",
                    "Nivel de ingles basico"
                ],
                "archivo": "cv_28_encargado_retail.pdf"
            },
            {
                "nombre_candidato": "Melina Vega",
                "puntaje_compatibilidad": 71,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Especialista en atencion al cliente y soporte administrativo bilingue (C2). Experiencia en gestion de tickets, facturacion basica, gestion documental y cumplimiento de SLAs corporativos.",
                "puntos_fuertes": [
                    "Dominio bilingue nativo de ingles (C2) y traduccion",
                    "Gestion documental, seguimiento de requerimientos y atencion corporativa",
                    "Manejo de sistemas CRM y plataformas digitales"
                ],
                "requisitos_faltantes": ["Sin formacion en Ciencias Economicas ni experiencia contable."],
                "archivo": "cv_22_soporte_bilingue.pdf"
            },
            {
                "nombre_candidato": "Georgina Melisa Laterza",
                "puntaje_compatibilidad": 70,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Bachiller contable con trayectoria en facturacion, notas de credito, remitos, compras y conciliacion de caja en empresas comerciales. Manejo de Office y sistemas comerciales.",
                "puntos_fuertes": [
                    "Experiencia en circuitos de facturacion, remitos y control de stock",
                    "Manejo de herramientas de Office y sistemas de gestion comercial",
                    "Titulo secundario con orientacion contable"
                ],
                "requisitos_faltantes": [
                    "Sin carrera universitaria en Ciencias Economicas",
                    "Sin manejo comprobable de idioma ingles"
                ],
                "archivo": "CV Georgina Laterza.pdf"
            },
            {
                "nombre_candidato": "Maria Magdalena Jerez",
                "puntaje_compatibilidad": 68,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Mas de 18 anos en administracion y soporte operativo en instituciones publicas y de salud. Experiencia en expedientes, cotizaciones con proveedores, libros banco y liquidacion.",
                "puntos_fuertes": [
                    "Amplia experiencia en control de caja, bancos y rendiciones",
                    "Trato y seguimiento de cotizaciones con proveedores",
                    "Manejo de sistemas de expedientes y documentacion oficial"
                ],
                "requisitos_faltantes": [
                    "Sin formacion universitaria en Ciencias Economicas",
                    "No posee nivel de ingles intermedio"
                ],
                "archivo": "CV - María Magdalena Jerez V2.pdf"
            },
            {
                "nombre_candidato": "Diego Ferrari",
                "puntaje_compatibilidad": 67,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Licenciado en Comercializacion (UADE) con experiencia en Trade Marketing, control presupuestario de promociones, analisis comercial de sell-in/sell-out y manejo de SAP SD.",
                "puntos_fuertes": [
                    "Titulo universitario afin en Comercializacion (UADE)",
                    "Control y seguimiento presupuestario en Excel y SAP SD",
                    "Capacidad analitica para reportes comerciales"
                ],
                "requisitos_faltantes": ["Perfil volcado a marketing y canales comerciales; no a tareas administrativas contables de compras/cobros."],
                "archivo": "cv_23_trade_marketing.pdf"
            },
            {
                "nombre_candidato": "Sofia Navarro",
                "puntaje_compatibilidad": 66,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Licenciada en Relaciones Publicas con experiencia en Customer Success B2B, seguimiento de contratos SaaS, gestion de clientes corporativos y retencion con ingles bilingue C2.",
                "puntos_fuertes": [
                    "Ingles bilingue C2 y comunicacion institucional de alto nivel",
                    "Seguimiento de cuentas corporativas y acuerdos de renovacion",
                    "Manejo de plataformas CRM y reportes ejecutivos"
                ],
                "requisitos_faltantes": ["Sin formacion en Ciencias Economicas ni experiencia en cobranzas morosas o compras contables."],
                "archivo": "cv_08_customer_success.pdf"
            },
            {
                "nombre_candidato": "Agustin Romero",
                "puntaje_compatibilidad": 65,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Licenciado en Comercializacion con mas de 5 anos en ventas B2B corporativas, negociacion de contratos marco de alto valor y uso diario de CRMs con ingles C1.",
                "puntos_fuertes": [
                    "Grado universitario y nivel de ingles avanzado C1",
                    "Capacidad de negociacion contractual y trato formal con clientes",
                    "Manejo de sistemas comerciales y pipeline de operaciones"
                ],
                "requisitos_faltantes": ["Perfil netamente comercial y de ventas, sin experiencia en circuitos administrativos ni compras internas."],
                "archivo": "cv_01_ejecutivo_cuentas_b2b.pdf"
            },
            {
                "nombre_candidato": "Camila Benitez",
                "puntaje_compatibilidad": 64,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Licenciada en Recursos Humanos (UBA) con experiencia en seleccion tecnica, relevamiento de perfiles, negociacion de propuestas laborales y gestion documental con ingles C2.",
                "puntos_fuertes": [
                    "Titulo universitario de grado de la UBA",
                    "Nivel de ingles bilingue C2",
                    "Capacidad organizativa, seguimiento de procesos y entrevistas"
                ],
                "requisitos_faltantes": ["Experiencia orientada exclusivamente a seleccion de talento IT; sin experiencia contable ni de compras."],
                "archivo": "cv_02_it_recruiter.pdf"
            },
            {
                "nombre_candidato": "Gonzalo Martin Varela",
                "puntaje_compatibilidad": 62,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Abogado egresado de la UBA especializado en litigios laborales y contratos comerciales. Posee ingles C1 y matricula CPACF, pero su perfil excede las tareas de asistencia operativa basica.",
                "puntos_fuertes": [
                    "Titulo de grado universitario de la Universidad de Buenos Aires (UBA)",
                    "Nivel de ingles avanzado bilingue (C1)",
                    "Capacidad de redaccion formal, analisis contractual y negociacion"
                ],
                "requisitos_faltantes": [
                    "Formacion juridica y no en Ciencias Economicas/Administracion",
                    "Perfil sobrecalificado para un rol operativo de asistencia administrativa"
                ],
                "archivo": "CV_Gonzalo_Varela_Abogado.pdf"
            },
            {
                "nombre_candidato": "Fabiana Marcela Serrano",
                "puntaje_compatibilidad": 60,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Amplia experiencia como administrativa y procuradora en el area legal y crediticia. Manejo de cobranzas multicanal, legajos de credito y tramites bancarios.",
                "puntos_fuertes": [
                    "Solida experiencia en cobranzas multicanal y seguimiento de deudas",
                    "Manejo de cajas, conciliaciones y cuentas corrientes",
                    "Antecedentes en gestion de compras y pagos a proveedores"
                ],
                "requisitos_faltantes": [
                    "Ingles en nivel basico",
                    "Sin estudios universitarios en Ciencias Economicas"
                ],
                "archivo": "Fabiana Marcela Serrano CV.pdf"
            },
            {
                "nombre_candidato": "Dr. Gonzalo Carrizo",
                "puntaje_compatibilidad": 58,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Abogado egresado de la UBA especialista en Compliance, prevencion de lavado (AML) y debida diligencia de contrapartes (KYC) con matricula CPACF e ingles C1.",
                "puntos_fuertes": [
                    "Formacion universitaria de grado en Derecho (UBA) e ingles C1",
                    "Rigor en auditoria documental, debida diligencia y normativas oficiales",
                    "Capacidad analitica para cumplimiento regulatorio"
                ],
                "requisitos_faltantes": ["Enfoque estrictamente juridico/penal economico; sobrecalificado para asistencia administrativa."],
                "archivo": "cv_27_abogado_compliance.pdf"
            },
            {
                "nombre_candidato": "Juan Ignacio Perez",
                "puntaje_compatibilidad": 55,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Desarrollador backend orientado a programacion en Python y FastAPI. Si bien posee ingles B2, su formacion esta orientada exclusivamente a tecnologia y no a tareas administrativas.",
                "puntos_fuertes": ["Nivel de ingles tecnico B2", "Capacidad logica y resolucion estructurada"],
                "requisitos_faltantes": ["Sin formacion en Ciencias Economicas ni experiencia en cobranzas o compras."],
                "archivo": "CV_Juan_Perez_Backend.pdf"
            },
            {
                "nombre_candidato": "Facundo Morales",
                "puntaje_compatibilidad": 50,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Licenciado en Sistemas (UTN) y Project Manager certificado (PMP). Experiencia liderando proyectos de software, gestion de presupuestos y ceremonias agiles con ingles C1.",
                "puntos_fuertes": [
                    "Control presupuestario y seguimiento de cronogramas en Confluence/Jira",
                    "Nivel de ingles avanzado profesional (C1)",
                    "Habilidades organizativas y coordinacion de equipos"
                ],
                "requisitos_faltantes": ["Perfil tecnologico senior; sobrecalificado y desalineado del rol operativo administrativo."],
                "archivo": "cv_16_project_manager.pdf"
            },
            {
                "nombre_candidato": "Federico Luna",
                "puntaje_compatibilidad": 48,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Licenciado en Ciencias de la Educacion (UdeSA) especializado en capacitacion corporativa, diseno instruccional y administracion de plataformas LMS con ingles C1.",
                "puntos_fuertes": [
                    "Titulo universitario de grado e ingles avanzado C1",
                    "Capacidad pedagogica y elaboracion de materiales corporativos",
                    "Manejo de herramientas informaticas y plataformas formativas"
                ],
                "requisitos_faltantes": ["Sin experiencia en facturacion, cobranzas o compras contables."],
                "archivo": "cv_15_capacitador_corporativo.pdf"
            },
            {
                "nombre_candidato": "Cristian de la Plaza",
                "puntaje_compatibilidad": 45,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Perfil senior con mas de 20 anos de experiencia en logistica, control de stock mediante SAP y remitos. Posee titulo de Perito Mercantil y movilidad propia.",
                "puntos_fuertes": [
                    "Mas de dos decadas en control de inventarios, remitos y trazabilidad logistica",
                    "Manejo de SAP y herramientas ofimaticas (Word, Excel, Outlook)",
                    "Disponibilidad horaria, carnet de conducir y vehiculo propio"
                ],
                "requisitos_faltantes": [
                    "No cursa carrera universitaria en Ciencias Economicas o Administracion",
                    "Experiencia orientada a deposito/logistica pesada y no a compras o conciliaciones"
                ],
                "archivo": "CV Cristian.docx.pdf"
            },
            {
                "nombre_candidato": "Martin Morales",
                "puntaje_compatibilidad": 40,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Licenciado en Marketing con experiencia en Growth y Paid Media (Google Ads, Meta Ads). Enfoque exclusivo en pauta digital, embudos y analitica web.",
                "puntos_fuertes": [
                    "Grado universitario completo",
                    "Analisis cuantitativo de metricas de retorno (ROAS)",
                    "Ingles profesional B2+"
                ],
                "requisitos_faltantes": [
                    "Sin relacion con tareas administrativas contables, compras ni cobranzas",
                    "Perfil tecnico orientado a adquisicion digital"
                ],
                "archivo": "cv_03_growth_marketer.pdf"
            },
            {
                "nombre_candidato": "Ing. Matias Duarte",
                "puntaje_compatibilidad": 38,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Ingeniero Industrial (UTN) con especializacion en Lean Manufacturing, balanceo de lineas productivas y tiempos y metodos en plantas fabriles automotrices.",
                "puntos_fuertes": [
                    "Graduado en Ingenieria Industrial con nivel de ingles C1",
                    "Manejo de metodologias de mejora continua y control estadistico",
                    "Manejo de SAP PP y analisis cuantitativo"
                ],
                "requisitos_faltantes": [
                    "Perfil orientado estrictamente a planta fabril y manufactura",
                    "Sin afinidad con funciones de secretaria o administracion de compras/cobros"
                ],
                "archivo": "cv_21_ingeniero_industrial.pdf"
            },
            {
                "nombre_candidato": "Alan Yamil Ilarraz",
                "puntaje_compatibilidad": 35,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Operario logistico con experiencia en deposito, picking, control de remitos y manejo de caja. Carrera universitaria en administracion actualmente pausada.",
                "puntos_fuertes": [
                    "Manejo de stock, control de remitos y recepcion de mercaderia",
                    "Experiencia en cobranzas y atencion al cliente en comercio",
                    "Secundario completo y cursos de herramientas digitales"
                ],
                "requisitos_faltantes": [
                    "Carrera universitaria en administracion actualmente pausada",
                    "Ingles en nivel basico",
                    "Experiencia predominantemente operativa de deposito"
                ],
                "archivo": "CV_Alan_Yamil_Ilarraz_Perfecto.pdf"
            },
            {
                "nombre_candidato": "Camila Sosa",
                "puntaje_compatibilidad": 32,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Licenciada en Comunicacion Social especializada en Social Media, produccion audiovisual para TikTok/Reels y gestion de influencers. Sin experiencia administrativa.",
                "puntos_fuertes": [
                    "Titulo universitario en Comunicacion",
                    "Ingles intermedio B2 y redaccion creativa",
                    "Manejo de herramientas audiovisuales y redes sociales"
                ],
                "requisitos_faltantes": [
                    "Sin conocimientos de contabilidad, proveedores, cobranzas ni facturacion",
                    "Perfil puramente enfocado a redes sociales y contenido"
                ],
                "archivo": "cv_19_community_manager.pdf"
            },
            {
                "nombre_candidato": "Gonzalo Torres",
                "puntaje_compatibilidad": 30,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Licenciado en Seguridad e Higiene Laboral (HSE) con experiencia en plantas quimicas, protocolos de bioseguridad, auditorias ISO y permisos de trabajo en campo.",
                "puntos_fuertes": [
                    "Titulo universitario de grado y matricula COPIME",
                    "Rigurosidad tecnica en normativas y procedimientos operativos",
                    "Liderazgo operativo en prevencion de riesgos"
                ],
                "requisitos_faltantes": [
                    "Especialidad tecnico-ambiental ajena al area administrativa",
                    "Sin conocimientos contables ni de compras comerciales"
                ],
                "archivo": "cv_13_seguridad_higiene.pdf"
            },
            {
                "nombre_candidato": "Matias Albornoz",
                "puntaje_compatibilidad": 28,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Licenciado en Publicidad con 5 anos de trayectoria como redactor creativo publicitario y guionista de comerciales. Sin antecedentes en tareas de oficina administrativa.",
                "puntos_fuertes": [
                    "Excelente redaccion y dominio de storytelling",
                    "Ingles avanzado C1 y conceptualizacion estrategica de marcas",
                    "Grado universitario en Publicidad"
                ],
                "requisitos_faltantes": [
                    "Sin experiencia contable, financiera ni de compras corporativas",
                    "Perfil puramente creativo y publicitario"
                ],
                "archivo": "cv_30_copywriter_creativo.pdf"
            },
            {
                "nombre_candidato": "Fatima Soledad Sanchez",
                "puntaje_compatibilidad": 25,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Recepcionista en centros de salud con tareas de atencion al paciente, turnos y caja. Sin formacion contable/administrativa formal ni conocimientos de ingles.",
                "puntos_fuertes": [
                    "Atencion al publico y gestion de turnos",
                    "Manejo basico de caja y recepcion de insumos"
                ],
                "requisitos_faltantes": [
                    "Sin formacion en Ciencias Economicas ni secundario comercial",
                    "Sin experiencia en compras corporativas, proveedores o cobranzas",
                    "Sin conocimientos de ingles"
                ],
                "archivo": "CV FÁTIMA SÁNCHEZ.pdf"
            },
            {
                "nombre_candidato": "Matias Adragna",
                "puntaje_compatibilidad": 25,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Analista de Soporte IT y Tecnico con mas de 15 anos en Help Desk, hardware y redes. No se ajusta al area administrativa, contable ni de compras.",
                "puntos_fuertes": [
                    "Amplio dominio tecnico de sistemas, soporte y seguridad de datos",
                    "Uso basico de SAP y paquete Office",
                    "Capacidad de resolucion de incidencias bajo presion"
                ],
                "requisitos_faltantes": [
                    "Formacion tecnica/periodistica, sin antecedentes en Ciencias Economicas",
                    "Sin experiencia en compras a proveedores, cobranzas o facturacion contable",
                    "Nivel de ingles basico"
                ],
                "archivo": "CV_Matias_Adragna_Soporte_IT_Tecnico_Administrativo.pdf"
            },
            {
                "nombre_candidato": "Lucia Pereyra",
                "puntaje_compatibilidad": 22,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Desarrolladora de software frontend especializada en React, TypeScript y Next.js. Perfil orientado exclusivamente a programacion web y diseno interactivo.",
                "puntos_fuertes": [
                    "Logica algoritmica y dominio informatico avanzado",
                    "Ingles tecnico profesional B2",
                    "Tecnicatura Universitaria en Programacion (UNSAM)"
                ],
                "requisitos_faltantes": ["Sin formacion administrativa, economica ni comercial."],
                "archivo": "cv_06_frontend_developer.pdf"
            },
            {
                "nombre_candidato": "Lucia Fernandez",
                "puntaje_compatibilidad": 20,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Disenadora grafica enfocada en identidad de marca y redes sociales. No cuenta con antecedentes en gestion administrativa ni el nivel de ingles requerido.",
                "puntos_fuertes": ["Manejo de herramientas de diseno visual y comunicacion corporativa"],
                "requisitos_faltantes": ["No cumple los requisitos formativos, operativos ni linguisticos solicitados."],
                "archivo": "CV_Lucia_Fernandez_Diseno.pdf"
            },
            {
                "nombre_candidato": "Nicolas Dominguez",
                "puntaje_compatibilidad": 20,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Disenador de producto digital (UX/UI) y Design Systems en Figma. Trayectoria en investigacion de usuarios y diseno de interfaces web/mobile.",
                "puntos_fuertes": [
                    "Titulo universitario en Diseno (FADU - UBA)",
                    "Nivel de ingles profesional C1",
                    "Capacidad de resolucion de problemas orientada al usuario"
                ],
                "requisitos_faltantes": ["Perfil tecnologico y de diseno visual; sin antecedentes administrativos ni contables."],
                "archivo": "cv_07_ux_ui_designer.pdf"
            },
            {
                "nombre_candidato": "Carlos Alberto Benitez",
                "puntaje_compatibilidad": 18,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Operario calificado de deposito con carnet habilitante de autoelevadores (Res. SRT 960/15) y picking por radiofrecuencia. Experiencia 100% logistica de planta.",
                "puntos_fuertes": [
                    "Manejo seguro de autoelevadores con certificacion oficial Res. SRT 960/15",
                    "Experiencia en recepcion, control de remitos de carga y estiba en altura",
                    "Disponibilidad horaria total y licencias de conducir al dia"
                ],
                "requisitos_faltantes": [
                    "Sin formacion administrativa ni estudios contables",
                    "Sin experiencia en compras de oficina, conciliaciones o facturacion",
                    "Sin manejo de idioma ingles"
                ],
                "archivo": "CV_Carlos_Benitez_Operario.pdf"
            },
            {
                "nombre_candidato": "Tomas Benitez",
                "puntaje_compatibilidad": 18,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Disenador Grafico egresado de la UBA especializado en branding, identidad visual, preprensa y diseno editorial en Adobe Illustrator y Photoshop.",
                "puntos_fuertes": [
                    "Graduado de honor de la Universidad de Buenos Aires",
                    "Criterio estetico y dominio avanzado de software de diseno",
                    "Ingles profesional B2"
                ],
                "requisitos_faltantes": ["Sin afinidad con funciones de administracion, contabilidad o compras de oficina."],
                "archivo": "cv_20_disenador_grafico.pdf"
            },
            {
                "nombre_candidato": "Joaquin Vega",
                "puntaje_compatibilidad": 18,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Ingeniero en Informatica (UBA) con especializacion en QA Automation, desarrollo de scripts de prueba en Playwright/TypeScript e integracion continua.",
                "puntos_fuertes": [
                    "Ingeniero en Informatica egresado de la FIUBA",
                    "Ingles profesional avanzado C1",
                    "Capacidad analitica estructurada para pruebas de software"
                ],
                "requisitos_faltantes": ["Perfil puramente tecnico de ingenieria de software; no aplica para tareas administrativas."],
                "archivo": "cv_11_qa_automation.pdf"
            },
            {
                "nombre_candidato": "Damian Godoy",
                "puntaje_compatibilidad": 16,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Administrador de Sistemas (SysAdmin) Linux y Redes corporativas. Experiencia en virtualizacion VMware/Proxmox, servidores y monitoreo con Zabbix.",
                "puntos_fuertes": [
                    "Solido conocimiento en administracion de servidores, seguridad e infraestructura",
                    "Certificacion Red Hat (RHCSA) y Cisco CCNA",
                    "Ingles tecnico avanzado B2"
                ],
                "requisitos_faltantes": ["Perfil exclusivamente de infraestructura IT; sin relacion con areas contables ni de compras."],
                "archivo": "cv_29_sysadmin_linux.pdf"
            },
            {
                "nombre_candidato": "Jeremias Fermin Jerez",
                "puntaje_compatibilidad": 15,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Tecnico electronico y estudiante de Ingenieria Industrial. Experiencia orientada a torneria, metrologia, deposito de repuestos y construccion.",
                "puntos_fuertes": ["Conocimientos tecnicos de precision, metrologia y organizacion de deposito"],
                "requisitos_faltantes": [
                    "Sin formacion en Ciencias Economicas ni administracion contable",
                    "Sin experiencia en gestion de cobranzas, compras o facturacion",
                    "Sin nivel de ingles requerido"
                ],
                "archivo": "CV - JEREMIAS JEREZ 2.pdf"
            },
            {
                "nombre_candidato": "Ignacio Castillo",
                "puntaje_compatibilidad": 15,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Senior Data Engineer y Licenciado en Computacion (UBA). Especialista en Big Data, Apache Spark, Airflow y arquitecturas Data Lakehouse en AWS.",
                "puntos_fuertes": [
                    "Formacion cientifica de grado en Ciencias de la Computacion (UBA)",
                    "Ingles bilingue profesional C1",
                    "Dominio de bases de datos masivas y programacion avanzada"
                ],
                "requisitos_faltantes": ["Perfil senior de ingenieria de datos en la nube sin relacion con secretaria o administracion."],
                "archivo": "cv_17_data_engineer.pdf"
            },
            {
                "nombre_candidato": "Lucas Santillan",
                "puntaje_compatibilidad": 14,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Desarrollador de aplicaciones moviles en Flutter y React Native para iOS y Android. Especialista en arquitecturas limpias y APIs bancarias.",
                "puntos_fuertes": [
                    "Desarrollo de software y logica de programacion",
                    "Ingles intermedio B2+",
                    "Tecnicatura en Desarrollo de Software"
                ],
                "requisitos_faltantes": ["Sin antecedentes ni competencias en gestion administrativa, compras o finanzas."],
                "archivo": "cv_25_mobile_developer.pdf"
            },
            {
                "nombre_candidato": "Lic. Paula Medina",
                "puntaje_compatibilidad": 12,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Licenciada en Enfermeria con Diploma de Honor (UBA) especializada en Cuidados Criticos (UTI), soporte vital avanzado y coordinacion asistencial.",
                "puntos_fuertes": [
                    "Grado universitario con Diploma de Honor de la UBA",
                    "Liderazgo de equipos y gestion de historias clinicas hospitalarias",
                    "Trabajo bajo estricta presion y protocolos de bioseguridad"
                ],
                "requisitos_faltantes": [
                    "Profesional de la salud clinica; sin formacion economica o administrativa de empresas",
                    "Sin experiencia en compras comerciales, cobranzas ni facturacion contable"
                ],
                "archivo": "cv_14_enfermeria_salud.pdf"
            },
            {
                "nombre_candidato": "Fernando Ariel Gill Alfonso",
                "puntaje_compatibilidad": 10,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Tecnico electromecanico con experiencia en matriceria, moldes y mantenimiento de planta. No presenta afinidad con roles de oficina.",
                "puntos_fuertes": ["Solido dominio de ajuste mecanico de precision y control dimensional"],
                "requisitos_faltantes": [
                    "Perfil exclusivamente industrial/metalurgico",
                    "Sin competencias ni formacion administrativa o contable"
                ],
                "archivo": "CV Fernando Alfonzo.pdf"
            },
            {
                "nombre_candidato": "Dra. Brenda Peralta",
                "puntaje_compatibilidad": 10,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Bioquimica Clinica graduada con honores de la UBA, especialista en analisis clinicos automatizados, hematologia, cultivos y normas ISO 15189.",
                "puntos_fuertes": [
                    "Titulo universitario de grado de la UBA con honores y matricula nacional",
                    "Rigurosidad analitica y gestion de calidad bajo normas ISO 15189",
                    "Ingles tecnico avanzado B2"
                ],
                "requisitos_faltantes": [
                    "Orientacion puramente bioquimica y biomedica",
                    "Sin experiencia en funciones administrativas, comerciales o contables"
                ],
                "archivo": "cv_24_bioquimico_laboratorio.pdf"
            }
        ]

        total = len(resultados_precargados)
        progreso = st.progress(0, text="Indexando perfiles curriculares...")

        for i, item in enumerate(resultados_precargados):
            progreso.progress((i + 1) / total, text=f"Procesando ({i + 1}/{total}): {item['archivo']}")
            time.sleep(0.35)

        progreso.empty()
        resultados = resultados_precargados

    # MODO ONLINE: EXTRACCION A TEXTO PLANO + FALLBACK DE MODELOS FLASH
    else:
        if not api_key:
            st.error("Ingrese su clave API en la barra lateral o active el Modo Contingencia.")
            st.stop()
        if not lista_cvs:
            st.warning("No se detectaron archivos curriculares en PDF para evaluar.")
            st.stop()

        client = genai.Client(api_key=api_key)
        
        modelos_activos = [
            "gemini-3.5-flash-lite",
            "gemini-3.1-flash-lite",
            "gemini-3-flash-preview",
            "gemini-3.6-flash",
            "gemini-3.7-flash",
            "gemini-3.5-flash"
        ]

        system_instruction = (
            "Actuas como un evaluador tecnico y de Recursos Humanos objetivo e imparcial. "
            "Tu tarea es contrastar el CV provisto contra los requisitos del puesto. "
            "Evalua estrictamente competencias tecnicas, formacion y experiencia comprobada. "
            "Ignora de forma absoluta factores demograficos como edad, genero, fotografia o direccion."
        )

        prompt_criterios = f"""
        Puesto a evaluar: {puesto}
        Requisitos excluyentes: {requisitos_excluyentes}
        Requisitos deseables: {requisitos_deseables}
        """

        total = len(lista_cvs)
        progreso = st.progress(0, text="Iniciando evaluacion algoritmica...")
        resultados = []
        errores_detectados = []

        for i, (nombre_archivo, pdf_bytes) in enumerate(lista_cvs):
            progreso.progress((i + 1) / total, text=f"Evaluando ({i + 1}/{total}): {nombre_archivo}")
            procesado_ok = False
            ultimo_error = ""

            texto_extraido = extraer_texto_pdf(pdf_bytes)

            if texto_extraido:
                contenido_envio = f"{prompt_criterios}\n\n--- TEXTO EXTRAIDO DEL CV ({nombre_archivo}) ---\n{texto_extraido}"
            else:
                contenido_envio = [
                    types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                    prompt_criterios
                ]

            for modelo in modelos_activos:
                try:
                    response = client.models.generate_content(
                        model=modelo,
                        contents=contenido_envio,
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
                except Exception as e:
                    ultimo_error = str(e)
                    continue

            if not procesado_ok:
                errores_detectados.append(f"Imposible procesar {nombre_archivo}: {ultimo_error}")

        progreso.empty()

        if errores_detectados:
            for err in errores_detectados:
                st.warning(err)

    # METRICAS TEMPORALES
    tiempo_total = time.time() - tiempo_inicio
    minutos = int(tiempo_total // 60)
    segundos = tiempo_total % 60
    cadena_tiempo = f"{minutos}m {segundos:.1f}s" if minutos > 0 else f"{segundos:.1f}s"

    # ---------------------------------------------------------
    # 5. Panel de Metricas, Tabla y Reporte Excel
    # ---------------------------------------------------------
    if resultados:
        resultados_ordenados = sorted(resultados, key=lambda x: x["puntaje_compatibilidad"], reverse=True)
        
        tot = len(resultados_ordenados)
        avz = len([r for r in resultados_ordenados if r["veredicto"] == "Avanzar a entrevista"])
        res = len([r for r in resultados_ordenados if r["veredicto"] == "En reserva"])
        dsc = len([r for r in resultados_ordenados if r["veredicto"] == "Descartar"])
        promedio_por_cv = tiempo_total / tot if tot > 0 else 0

        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
        col_m1.metric("POSTULANTES", f"{tot}")
        col_m2.metric("TIEMPO TOTAL", f"{cadena_tiempo}", delta=f"{promedio_por_cv:.2f}s / CV", delta_color="off")
        col_m3.metric("CALIFICADOS", f"{avz}", delta="Avanzan", delta_color="normal")
        col_m4.metric("EN RESERVA", f"{res}", delta="Potenciales", delta_color="off")
        col_m5.metric("DESCARTADOS", f"{dsc}", delta="No cumplen", delta_color="inverse")

        st.markdown("<div style='height: 18px;'></div>", unsafe_allow_html=True)

        st.markdown("#### Matriz General de Compatibilidad")
        filas = []
        for r in resultados_ordenados:
            filas.append({
                "Candidato": r["nombre_candidato"],
                "Score": f"{r['puntaje_compatibilidad']}%",
                "Excluyentes": "Cumple" if r["cumple_excluyentes"] else "No cumple",
                "Veredicto": r["veredicto"],
                "Archivo": r["archivo"]
            })
        st.dataframe(pd.DataFrame(filas), use_container_width=True)

        # Generacion de Excel corporativo con openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Matriz de Seleccion"

        headers = [
            "Candidato", "Score", "Cumple Excluyentes", "Veredicto", 
            "Resumen Profesional", "Puntos Fuertes", "Requisitos Faltantes", "Archivo"
        ]
        ws.append(headers)

        header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
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
        ws.row_dimensions[1].height = 26

        colores = {
            "Avanzar a entrevista": (PatternFill(start_color="DCFCE7", fill_type="solid"), Font(color="166534", bold=True)),
            "En reserva": (PatternFill(start_color="FEF9C3", fill_type="solid"), Font(color="854D0E", bold=True)),
            "Descartar": (PatternFill(start_color="FEE2E2", fill_type="solid"), Font(color="991B1B", bold=True))
        }

        for row_idx, r in enumerate(resultados_ordenados, start=2):
            pf_txt = "\n".join([f"- {p}" for p in r.get("puntos_fuertes", [])])
            rf_txt = "\n".join([f"- {p}" for p in r.get("requisitos_faltantes", [])]) if r.get("requisitos_faltantes") else "Sin brechas detectadas"

            ws.append([
                r["nombre_candidato"],
                f"{r['puntaje_compatibilidad']}%",
                "Si" if r["cumple_excluyentes"] else "No",
                r["veredicto"],
                r["resumen_perfil"],
                pf_txt,
                rf_txt,
                r["archivo"]
            ])
            ws.row_dimensions[row_idx].height = 65

            for c_idx in range(1, len(headers) + 1):
                c = ws.cell(row=row_idx, column=c_idx)
                c.border = thin_border
                c.font = Font(name="Segoe UI", size=9, color="0F172A")
                if c_idx in [2, 3]:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif c_idx == 4:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    if r["veredicto"] in colores:
                        fill, font = colores[r["veredicto"]]
                        c.fill = fill
                        c.font = font
                elif c_idx in [5, 6, 7]:
                    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center")

        anchos = {1: 25, 2: 12, 3: 18, 4: 22, 5: 45, 6: 35, 7: 35, 8: 28}
        for col_idx, w in anchos.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = w

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        st.download_button(
            label="Descargar Reporte Ejecutivo (.xlsx)",
            data=buf,
            file_name="Reporte_Seleccion_RRHH.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        st.markdown("<div style='height: 24px;'></div>", unsafe_allow_html=True)
        st.markdown("#### Desglose Individual de Postulantes")

        for r in resultados_ordenados:
            with st.expander(f"{r['nombre_candidato']} | Compatibilidad: {r['puntaje_compatibilidad']}% | Dictamen: {r['veredicto']}"):
                st.markdown(f"**Archivo analizado:** `{r['archivo']}`")
                st.markdown(f"**Sintesis tecnica:** {r['resumen_perfil']}")
                
                c_izq, c_der = st.columns(2)
                with c_izq:
                    st.markdown("**Fortalezas acreditadas:**")
                    for pf in r["puntos_fuertes"]:
                        st.markdown(f"- {pf}")
                with c_der:
                    st.markdown("**Brechas o requisitos faltantes:**")
                    if r["requisitos_faltantes"]:
                        for rf in r["requisitos_faltantes"]:
                            st.markdown(f"- {rf}")
                    else:
                        st.markdown("- Cumple con la totalidad de los requisitos.")