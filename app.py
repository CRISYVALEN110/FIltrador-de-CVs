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

    # RAMA 1: MODO SIMULADO / OFFLINE (Base consolidada de 55 candidatos)
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
                "nombre_candidato": "Mariana Rivas",
                "puntaje_compatibilidad": 95,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Asistente ejecutiva bilingüe (C2) con más de 7 años asistiendo a directores y gerencias. Sólida experiencia en organización administrativa, rendición de gastos, minutas, gestión de compras de servicios y soporte operativo integral.",
                "puntos_fuertes": [
                    "Inglés bilingüe C2 acreditado para trato corporativo",
                    "Amplia experiencia en soporte ejecutivo, rendición de gastos y compras",
                    "Dominio avanzado de Google Workspace, Microsoft 365 y ERPs"
                ],
                "requisitos_faltantes": ["Estudios orientados a Relaciones Institucionales y no estrictamente a Ciencias Económicas."],
                "archivo": "cv_26_asistente_ejecutiva.pdf"
            },
            {
                "nombre_candidato": "Patricia Mónica Figari",
                "puntaje_compatibilidad": 94,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Estudiante avanzada de Contador Público (UBA - 26 materias aprobadas) con más de 15 años de experiencia liderando cuentas a pagar, gestión de proveedores, cobranzas y conciliaciones contables. Cuenta con manejo de Bejerman, Tango y BAS.",
                "puntos_fuertes": [
                    "Formación universitaria en Ciencias Económicas (UBA)",
                    "Dominio integral de cuentas a pagar, compras a proveedores y cobranzas",
                    "Manejo de sistemas ERP (Tango, Bejerman, BAS) y herramientas de IA aplicada"
                ],
                "requisitos_faltantes": ["Acreditación formal de nivel de inglés intermedio en CV."],
                "archivo": "Patricia Mónica Figari 05 2026 CV.pdf"
            },
            {
                "nombre_candidato": "Mariana Silva",
                "puntaje_compatibilidad": 93,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Licenciada en Administración con más de 5 años gestionando compras productivas e indirectas, negociación con proveedores, emisión de órdenes de compra en SAP MM y análisis exhaustivo de costos.",
                "puntos_fuertes": [
                    "Título universitario en Administración (UADE)",
                    "Especialista en compras corporativas, licitaciones y gestión de proveedores",
                    "Manejo avanzado de SAP MM, Excel y herramientas de abastecimiento"
                ],
                "requisitos_faltantes": ["Perfil fuertemente enfocado en compras; requerirá inducción en tareas directas de cobranzas."],
                "archivo": "cv_12_comprador_procurement.pdf"
            },
            {
                "nombre_candidato": "Vanina Giselle Herrera",
                "puntaje_compatibilidad": 92,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Estudiante universitaria en la UBA (Lic. en Relaciones del Trabajo) con inglés intermedio certificado. Amplia trayectoria en compras de insumos, pago a proveedores, gestión de cobranzas, trámites bancarios y cash flow.",
                "puntos_fuertes": [
                    "Carrera universitaria afín en curso en UBA y diplomatura",
                    "Experiencia comprobada en cobranzas, proveedores, facturación y bancos",
                    "Nivel de inglés intermedio (Oxford University English Course)"
                ],
                "requisitos_faltantes": ["Mayor profundización en herramientas de IA generativa."],
                "archivo": "Cv Vanina Herrera.pdf"
            },
            {
                "nombre_candidato": "Mariana Sofía Rossi",
                "puntaje_compatibilidad": 91,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Técnica Superior en Administración con más de 5 años en facturación masiva ARCA, cuentas corrientes, proveedores, conciliaciones bancarias y ERP Tango. Sólida experiencia contable y manejo fluido de Excel.",
                "puntos_fuertes": [
                    "Formación técnica acreditada en Administración General",
                    "Experiencia en facturación masiva ARCA, cuentas corrientes y proveedores",
                    "Dominio de sistemas de gestión (Tango Gestión y Bejerman)"
                ],
                "requisitos_faltantes": ["No posee certificación de nivel de inglés intermedio formal."],
                "archivo": "CV_Mariana_Rossi_Administrativa.pdf"
            },
            {
                "nombre_candidato": "Esteban Castro",
                "puntaje_compatibilidad": 90,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Contador Público Nacional con más de 6 años en administración de personal, liquidación salarial, Libro Sueldos Digital AFIP, conciliaciones bancarias y manejo avanzado de sistemas ERP (Tango Sueldos, SAP).",
                "puntos_fuertes": [
                    "Título universitario de Contador Público (UNLZ)",
                    "Dominio profundo de normativa contable, impositiva (AFIP) y convenios colectivos",
                    "Manejo de sistemas Tango, SAP y conciliaciones complejas"
                ],
                "requisitos_faltantes": ["Nivel de inglés técnico básico."],
                "archivo": "cv_09_hard_hr_payroll.pdf"
            },
            {
                "nombre_candidato": "Florencia Herrera",
                "puntaje_compatibilidad": 89,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Licenciada en Economía (UTDT) con 5 años en análisis financiero, cash flow proyectado, control presupuestario y manejo de SAP S/4HANA. Cuenta con inglés bilingüe (C1) y dominio avanzado de Excel.",
                "puntos_fuertes": [
                    "Formación universitaria de grado en Ciencias Económicas (Lic. en Economía)",
                    "Inglés avanzado bilingüe C1 y modelado financiero en Excel/Power BI",
                    "Dominio de ERP SAP S/4HANA y conciliación presupuestaria"
                ],
                "requisitos_faltantes": ["Perfil orientado a análisis estratégico financiero más que a trámites operativos."],
                "archivo": "cv_04_analista_fpa.pdf"
            },
            {
                "nombre_candidato": "Evelyn Gisela Maturano",
                "puntaje_compatibilidad": 88,
                "cumple_excluyentes": False,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Sólido perfil administrativo-financiero con experiencia en compras estratégicas, proveedores, cobranzas complejas (aging), conciliaciones y módulo MIRO de SAP. Posee inglés intermedio certificado.",
                "puntos_fuertes": [
                    "Amplio dominio de SAP, Tango, Catedral y Excel avanzado",
                    "Gestión integral de compras, cuentas corrientes, pagos y cobranzas",
                    "Inglés intermedio certificado y cursos impositivos en la UBA"
                ],
                "requisitos_faltantes": ["No cursa actualmente carrera de grado en Ciencias Económicas."],
                "archivo": "CV_Evelyn_Maturano (2).pdf"
            },
            {
                "nombre_candidato": "Valeria Giménez",
                "puntaje_compatibilidad": 87,
                "cumple_excluyentes": True,
                "veredicto": "Avanzar a entrevista",
                "resumen_perfil": "Licenciada en Finanzas con 5 años evaluando riesgo crediticio, gestión de cobranzas de cuentas en mora, análisis de estados contables de PyMEs y seguimiento de flujos de caja en instituciones financieras.",
                "puntos_fuertes": [
                    "Grado universitario en Finanzas (UADE)",
                    "Sólida experiencia en cobranzas, análisis crediticio y solvencia de clientes",
                    "Inglés intermedio B2 y manejo avanzado de Excel y herramientas contables"
                ],
                "requisitos_faltantes": ["Experiencia en compras operativas menor que en cobranzas y análisis financiero."],
                "archivo": "cv_18_analista_riesgo_credito.pdf"
            },
            {
                "nombre_candidato": "Juan Manuel Mansilla",
                "puntaje_compatibilidad": 84,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Administrativo generalista con 9 años de experiencia en Pymes, destacándose en compras operativas, gestión de proveedores, conciliación de facturación con ERP y cobranzas.",
                "puntos_fuertes": [
                    "Especialista en compras Pyme, proveedores y conciliación documental",
                    "Manejo de sistemas ERP, facturación ARCA y arqueos de caja",
                    "Bachiller en Economía y Administración"
                ],
                "requisitos_faltantes": [
                    "No cursa carrera de grado en Ciencias Económicas",
                    "Sin acreditación de nivel de inglés intermedio"
                ],
                "archivo": "CV Juan Manuel Mansilla - Adm.pdf"
            },
            {
                "nombre_candidato": "Lucas Gabriel Medina",
                "puntaje_compatibilidad": 83,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Técnico en RRHH con 4 años en control de ausentismo, novedades de liquidación, legajos y contacto con ART. Perfil administrativo orientado a gestión de personal y trámites.",
                "puntos_fuertes": [
                    "Experiencia en control de novedades y administración de personal",
                    "Manejo de sistemas de fichaje y plataformas oficiales (ARCA/ART)",
                    "Tecnicatura Superior en Recursos Humanos finalizada"
                ],
                "requisitos_faltantes": [
                    "Experiencia orientada a personal y no a compras comerciales o cobranzas",
                    "Nivel de inglés básico"
                ],
                "archivo": "CV_Lucas_Medina_AdmRRHH.pdf"
            },
            {
                "nombre_candidato": "Esteban Benitez",
                "puntaje_compatibilidad": 82,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Experiencia destacada en cobranzas, tesorería, control de recaudación y cuentas corrientes en retail. Manejo avanzado de Excel aplicado a Cash Flow con secundario contable.",
                "puntos_fuertes": [
                    "Dominio de cobranzas, medios de pago y tesorería",
                    "Excel avanzado aplicado a Cash Flow y conciliaciones",
                    "Inglés en curso en Liceo Cultural Británico"
                ],
                "requisitos_faltantes": ["Sin formación universitaria en curso en Ciencias Económicas."],
                "archivo": "CV BENITEZ ESTEBAN.pdf"
            },
            {
                "nombre_candidato": "Carla Agostina Antognoli",
                "puntaje_compatibilidad": 80,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Administrativa con 6 años de experiencia en soporte contable, ventas mayoristas, seguimiento de cuentas corrientes, facturación, echeq y conciliaciones bancarias en SAP y Tango.",
                "puntos_fuertes": [
                    "Manejo de sistemas de gestión (SAP, Tango) y Excel",
                    "Experiencia en conciliaciones bancarias, cuentas corrientes y facturación",
                    "Nivel de inglés intermedio y movilidad propia"
                ],
                "requisitos_faltantes": ["No cursa estudios de grado en Ciencias Económicas."],
                "archivo": "CVCarla_Agostina_Antognoli.pdf"
            },
            {
                "nombre_candidato": "Julieta Díaz",
                "puntaje_compatibilidad": 79,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Licenciada en Comercio Internacional con 5 años coordinando compras y logística de importación, control de documentación comercial, facturación y pagos bancarios al exterior con inglés C1.",
                "puntos_fuertes": [
                    "Experiencia en gestión documental de compras internacionales y proveedores",
                    "Nivel de inglés avanzado bilingüe (C1)",
                    "Manejo de SAP y trámites bancarios/comerciales"
                ],
                "requisitos_faltantes": ["Especialización en comercio exterior; requerirá adaptación al circuito administrativo local."],
                "archivo": "cv_10_comercio_exterior.pdf"
            },
            {
                "nombre_candidato": "Bárbara Hipler",
                "puntaje_compatibilidad": 78,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Graduada en Secretariado Administrativo y Perito Mercantil con orientación contable. Experiencia en compras de insumos, pago a proveedores, facturación y caja chica.",
                "puntos_fuertes": [
                    "Formación específica en Secretariado Administrativo y Perito Mercantil",
                    "Experiencia operativa en compras de insumos y pago a proveedores",
                    "Manejo de caja chica, trámites bancarios y facturación"
                ],
                "requisitos_faltantes": [
                    "Sin carrera de grado universitaria en Ciencias Económicas",
                    "Nivel de inglés sin certificación intermedia"
                ],
                "archivo": "CV Bárbara Hipler.pdf"
            },
            {
                "nombre_candidato": "Antonella Sol Falsetti",
                "puntaje_compatibilidad": 76,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Profesional con más de 6 años de experiencia en gestión documental, compras de insumos, control de stock y cobranzas corporativas con inglés intermedio.",
                "puntos_fuertes": [
                    "Experiencia sólida en compras de insumos, stock y cobranzas comerciales",
                    "Gestión documental, auditorías y sistemas integrados de gestión",
                    "Nivel de inglés intermedio acreditado"
                ],
                "requisitos_faltantes": ["Formación universitaria orientada a Higiene y Seguridad, no a Ciencias Económicas."],
                "archivo": "CV Antonella Falsetti.pdf"
            },
            {
                "nombre_candidato": "María Cecilia Martín",
                "puntaje_compatibilidad": 75,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Asistente administrativa bilingüe certificada (First Certificate in English). Experiencia en reclamo de cobranzas, gestión de cuentas corrientes, facturación y data entry.",
                "puntos_fuertes": [
                    "Certificación First Certificate in English (FCE)",
                    "Experiencia en gestión comercial de cuentas corrientes y cobranzas",
                    "Habilidades destacadas en redacción corporativa y precisión en carga de datos"
                ],
                "requisitos_faltantes": ["Sin formación académica en Ciencias Económicas o Administración."],
                "archivo": "María Cecilia Martín CV.pdf"
            },
            {
                "nombre_candidato": "Santiago Rossi",
                "puntaje_compatibilidad": 74,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Licenciado en Sistemas de Información de las Organizaciones (FCE-UBA) con experiencia en automatización de reportes, conciliación de datos en SQL y tableros en Power BI. Inglés C1.",
                "puntos_fuertes": [
                    "Egresado de la Facultad de Ciencias Económicas (FCE - UBA)",
                    "Nivel de inglés profesional avanzado (C1)",
                    "Capacidad analítica para procesamiento de datos, Excel y conciliaciones"
                ],
                "requisitos_faltantes": ["Perfil enfocado en Business Intelligence y analítica de datos, no en tareas operativas de oficina."],
                "archivo": "cv_05_data_analyst.pdf"
            },
            {
                "nombre_candidato": "Maximiliano Prado",
                "puntaje_compatibilidad": 72,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Encargado comercial con más de 6 años liderando sucursales de retail. Amplia práctica en arqueos de caja, facturación electrónica A/B, control de inventarios y rendición de valores.",
                "puntos_fuertes": [
                    "Dominio de arqueo de caja, facturación electrónica y medios de pago",
                    "Control riguroso de inventarios y rendición de caudales",
                    "Tecnicatura Universitaria en Gestión Retail (UNSAM)"
                ],
                "requisitos_faltantes": [
                    "Experiencia orientada al salón de ventas comercial y no a la administración corporativa",
                    "Nivel de inglés básico"
                ],
                "archivo": "cv_28_encargado_retail.pdf"
            },
            {
                "nombre_candidato": "Melina Vega",
                "puntaje_compatibilidad": 71,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Especialista en atención al cliente y soporte administrativo bilingüe (C2). Experiencia en gestión de tickets, facturación básica, gestión documental y cumplimiento de SLAs corporativos.",
                "puntos_fuertes": [
                    "Dominio bilingüe nativo de inglés (C2) y traducción",
                    "Gestión documental, seguimiento de requerimientos y atención corporativa",
                    "Manejo de sistemas CRM y plataformas digitales"
                ],
                "requisitos_faltantes": ["Sin formación en Ciencias Económicas ni experiencia contable."],
                "archivo": "cv_22_soporte_bilingue.pdf"
            },
            {
                "nombre_candidato": "Georgina Melisa Laterza",
                "puntaje_compatibilidad": 70,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Bachiller contable con trayectoria en facturación, notas de crédito, remitos, compras y conciliación de caja en empresas comerciales. Manejo de Office y sistemas comerciales.",
                "puntos_fuertes": [
                    "Experiencia en circuitos de facturación, remitos y control de stock",
                    "Manejo de herramientas de Office y sistemas de gestión comercial",
                    "Título secundario con orientación contable"
                ],
                "requisitos_faltantes": [
                    "Sin carrera universitaria en Ciencias Económicas",
                    "Sin manejo comprobable de idioma inglés"
                ],
                "archivo": "CV Georgina Laterza.pdf"
            },
            {
                "nombre_candidato": "María Magdalena Jerez",
                "puntaje_compatibilidad": 68,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Más de 18 años en administración y soporte operativo en instituciones públicas y de salud. Experiencia en expedientes, cotizaciones con proveedores, libros banco y liquidación.",
                "puntos_fuertes": [
                    "Amplia experiencia en control de caja, bancos y rendiciones",
                    "Trato y seguimiento de cotizaciones con proveedores",
                    "Manejo de sistemas de expedientes y documentación oficial"
                ],
                "requisitos_faltantes": [
                    "Sin formación universitaria en Ciencias Económicas",
                    "No posee nivel de inglés intermedio"
                ],
                "archivo": "CV - María Magdalena Jerez V2.pdf"
            },
            {
                "nombre_candidato": "Diego Ferrari",
                "puntaje_compatibilidad": 67,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Licenciado en Comercialización (UADE) con experiencia en Trade Marketing, control presupuestario de promociones, análisis comercial de sell-in/sell-out y manejo de SAP SD.",
                "puntos_fuertes": [
                    "Título universitario afín en Comercialización (UADE)",
                    "Control y seguimiento presupuestario en Excel y SAP SD",
                    "Capacidad analítica para reportes comerciales"
                ],
                "requisitos_faltantes": ["Perfil volcado a marketing y canales comerciales; no a tareas administrativas contables de compras/cobros."],
                "archivo": "cv_23_trade_marketing.pdf"
            },
            {
                "nombre_candidato": "Sofía Navarro",
                "puntaje_compatibilidad": 66,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Licenciada en Relaciones Públicas con experiencia en Customer Success B2B, seguimiento de contratos SaaS, gestión de clientes corporativos y retención con inglés bilingüe C2.",
                "puntos_fuertes": [
                    "Inglés bilingüe C2 y comunicación institucional de alto nivel",
                    "Seguimiento de cuentas corporativas y acuerdos de renovación",
                    "Manejo de plataformas CRM y reportes ejecutivos"
                ],
                "requisitos_faltantes": ["Sin formación en Ciencias Económicas ni experiencia en cobranzas morosas o compras contables."],
                "archivo": "cv_08_customer_success.pdf"
            },
            {
                "nombre_candidato": "Agustín Romero",
                "puntaje_compatibilidad": 65,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Licenciado en Comercialización con más de 5 años en ventas B2B corporativas, negociación de contratos marco de alto valor y uso diario de CRMs con inglés C1.",
                "puntos_fuertes": [
                    "Grado universitario y nivel de inglés avanzado C1",
                    "Capacidad de negociación contractual y trato formal con clientes",
                    "Manejo de sistemas comerciales y pipeline de operaciones"
                ],
                "requisitos_faltantes": ["Perfil netamente comercial y de ventas, sin experiencia en circuitos administrativos ni compras internas."],
                "archivo": "cv_01_ejecutivo_cuentas_b2b.pdf"
            },
            {
                "nombre_candidato": "Camila Benítez",
                "puntaje_compatibilidad": 64,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Licenciada en Recursos Humanos (UBA) con experiencia en selección técnica, relevamiento de perfiles, negociación de propuestas laborales y gestión documental con inglés C2.",
                "puntos_fuertes": [
                    "Título universitario de grado de la UBA",
                    "Nivel de inglés bilingüe C2",
                    "Capacidad organizativa, seguimiento de procesos y entrevistas"
                ],
                "requisitos_faltantes": ["Experiencia orientada exclusivamente a selección de talento IT; sin experiencia contable ni de compras."],
                "archivo": "cv_02_it_recruiter.pdf"
            },
            {
                "nombre_candidato": "Gonzalo Martín Varela",
                "puntaje_compatibilidad": 62,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Abogado egresado de la UBA especializado en litigios laborales y contratos comerciales. Posee inglés C1 y matrícula CPACF, pero su perfil excede las tareas de asistencia operativa básica.",
                "puntos_fuertes": [
                    "Título de grado universitario de la Universidad de Buenos Aires (UBA)",
                    "Nivel de inglés avanzado bilingüe (C1)",
                    "Capacidad de redacción formal, análisis contractual y negociación"
                ],
                "requisitos_faltantes": [
                    "Formación jurídica y no en Ciencias Económicas/Administración",
                    "Perfil sobrecalificado para un rol operativo de asistencia administrativa"
                ],
                "archivo": "CV_Gonzalo_Varela_Abogado.pdf"
            },
            {
                "nombre_candidato": "Fabiana Marcela Serrano",
                "puntaje_compatibilidad": 60,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Amplia experiencia como administrativa y procuradora en el área legal y crediticia. Manejo de cobranzas multicanal, legajos de crédito y trámites bancarios.",
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
                "nombre_candidato": "Dr. Gonzalo Carrizo",
                "puntaje_compatibilidad": 58,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Abogado egresado de la UBA especialista en Compliance, prevención de lavado (AML) y debida diligencia de contrapartes (KYC) con matrícula CPACF e inglés C1.",
                "puntos_fuertes": [
                    "Formación universitaria de grado en Derecho (UBA) e inglés C1",
                    "Rigor en auditoría documental, debida diligencia y normativas oficiales",
                    "Capacidad analítica para cumplimiento regulatorio"
                ],
                "requisitos_faltantes": ["Enfoque estrictamente jurídico/penal económico; sobrecalificado para asistencia administrativa."],
                "archivo": "cv_27_abogado_compliance.pdf"
            },
            {
                "nombre_candidato": "Juan Ignacio Pérez",
                "puntaje_compatibilidad": 55,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Desarrollador backend orientado a programación en Python y FastAPI. Si bien posee inglés B2, su formación está orientada exclusivamente a tecnología y no a tareas administrativas.",
                "puntos_fuertes": ["Nivel de inglés técnico B2", "Capacidad lógica y resolución estructurada"],
                "requisitos_faltantes": ["Sin formación en Ciencias Económicas ni experiencia en cobranzas o compras."],
                "archivo": "CV_Juan_Perez_Backend.pdf"
            },
            {
                "nombre_candidato": "Facundo Morales",
                "puntaje_compatibilidad": 50,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Licenciado en Sistemas (UTN) y Project Manager certificado (PMP). Experiencia liderando proyectos de software, gestión de presupuestos y ceremonias ágiles con inglés C1.",
                "puntos_fuertes": [
                    "Control presupuestario y seguimiento de cronogramas en Confluence/Jira",
                    "Nivel de inglés avanzado profesional (C1)",
                    "Habilidades organizativas y coordinación de equipos"
                ],
                "requisitos_faltantes": ["Perfil tecnológico sénior; sobrecalificado y desalineado del rol operativo administrativo."],
                "archivo": "cv_16_project_manager.pdf"
            },
            {
                "nombre_candidato": "Federico Luna",
                "puntaje_compatibilidad": 48,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Licenciado en Ciencias de la Educación (UdeSA) especializado en capacitación corporativa, diseño instruccional y administración de plataformas LMS con inglés C1.",
                "puntos_fuertes": [
                    "Título universitario de grado e inglés avanzado C1",
                    "Capacidad pedagógica y elaboración de materiales corporativos",
                    "Manejo de herramientas informáticas y plataformas formativas"
                ],
                "requisitos_faltantes": ["Sin experiencia en facturación, cobranzas o compras contables."],
                "archivo": "cv_15_capacitador_corporativo.pdf"
            },
            {
                "nombre_candidato": "Cristian de la Plaza",
                "puntaje_compatibilidad": 45,
                "cumple_excluyentes": False,
                "veredicto": "En reserva",
                "resumen_perfil": "Perfil senior con más de 20 años de experiencia en logística, control de stock mediante SAP y remitos. Posee título de Perito Mercantil y movilidad propia.",
                "puntos_fuertes": [
                    "Más de dos décadas en control de inventarios, remitos y trazabilidad logística",
                    "Manejo de SAP y herramientas ofimáticas (Word, Excel, Outlook)",
                    "Disponibilidad horaria, carnet de conducir y vehículo propio"
                ],
                "requisitos_faltantes": [
                    "No cursa carrera universitaria en Ciencias Económicas o Administración",
                    "Experiencia orientada a depósito/logística pesada y no a compras o conciliaciones"
                ],
                "archivo": "CV Cristian.docx.pdf"
            },
            {
                "nombre_candidato": "Martín Morales",
                "puntaje_compatibilidad": 40,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Licenciado en Marketing con experiencia en Growth y Paid Media (Google Ads, Meta Ads). Enfoque exclusivo en pauta digital, embudos y analítica web.",
                "puntos_fuertes": [
                    "Grado universitario completo",
                    "Análisis cuantitativo de métricas de retorno (ROAS)",
                    "Inglés profesional B2+"
                ],
                "requisitos_faltantes": [
                    "Sin relación con tareas administrativas contables, compras ni cobranzas",
                    "Perfil técnico orientado a adquisición digital"
                ],
                "archivo": "cv_03_growth_marketer.pdf"
            },
            {
                "nombre_candidato": "Ing. Matías Duarte",
                "puntaje_compatibilidad": 38,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Ingeniero Industrial (UTN) con especialización en Lean Manufacturing, balanceo de líneas productivas y tiempos y métodos en plantas fabriles automotrices.",
                "puntos_fuertes": [
                    "Graduado en Ingeniería Industrial con nivel de inglés C1",
                    "Manejo de metodologías de mejora continua y control estadístico",
                    "Manejo de SAP PP y análisis cuantitativo"
                ],
                "requisitos_faltantes": [
                    "Perfil orientado estrictamente a planta fabril y manufactura",
                    "Sin afinidad con funciones de secretaría o administración de compras/cobros"
                ],
                "archivo": "cv_21_ingeniero_industrial.pdf"
            },
            {
                "nombre_candidato": "Alan Yamil Ilarraz",
                "puntaje_compatibilidad": 35,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Operario logístico con experiencia en depósito, picking, control de remitos y manejo de caja. Carrera universitaria en administración actualmente pausada.",
                "puntos_fuertes": [
                    "Manejo de stock, control de remitos y recepción de mercadería",
                    "Experiencia en cobranzas y atención al cliente en comercio",
                    "Secundario completo y cursos de herramientas digitales"
                ],
                "requisitos_faltantes": [
                    "Carrera universitaria en administración actualmente pausada",
                    "Inglés en nivel básico",
                    "Experiencia predominantemente operativa de depósito"
                ],
                "archivo": "CV_Alan_Yamil_Ilarraz_Perfecto.pdf"
            },
            {
                "nombre_candidato": "Camila Sosa",
                "puntaje_compatibilidad": 32,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Licenciada en Comunicación Social especializada en Social Media, producción audiovisual para TikTok/Reels y gestión de influencers. Sin experiencia administrativa.",
                "puntos_fuertes": [
                    "Título universitario en Comunicación",
                    "Inglés intermedio B2 y redacción creativa",
                    "Manejo de herramientas audiovisuales y redes sociales"
                ],
                "requisitos_faltantes": [
                    "Sin conocimientos de contabilidad, proveedores, cobranzas ni facturación",
                    "Perfil puramente enfocado a redes sociales y contenido"
                ],
                "archivo": "cv_19_community_manager.pdf"
            },
            {
                "nombre_candidato": "Gonzalo Torres",
                "puntaje_compatibilidad": 30,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Licenciado en Seguridad e Higiene Laboral (HSE) con experiencia en plantas químicas, protocolos de bioseguridad, auditorías ISO y permisos de trabajo en campo.",
                "puntos_fuertes": [
                    "Título universitario de grado y matrícula COPIME",
                    "Rigurosidad técnica en normativas y procedimientos operativos",
                    "Liderazgo operativo en prevención de riesgos"
                ],
                "requisitos_faltantes": [
                    "Especialidad técnico-ambiental ajena al área administrativa",
                    "Sin conocimientos contables ni de compras comerciales"
                ],
                "archivo": "cv_13_seguridad_higiene.pdf"
            },
            {
                "nombre_candidato": "Matías Albornoz",
                "puntaje_compatibilidad": 28,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Licenciado en Publicidad con 5 años de trayectoria como redactor creativo publicitario y guionista de comerciales. Sin antecedentes en tareas de oficina administrativa.",
                "puntos_fuertes": [
                    "Excelente redacción y dominio de storytelling",
                    "Inglés avanzado C1 y conceptualización estratégica de marcas",
                    "Grado universitario en Publicidad"
                ],
                "requisitos_faltantes": [
                    "Sin experiencia contable, financiera ni de compras corporativas",
                    "Perfil puramente creativo y publicitario"
                ],
                "archivo": "cv_30_copywriter_creativo.pdf"
            },
            {
                "nombre_candidato": "Fátima Soledad Sánchez",
                "puntaje_compatibilidad": 25,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Recepcionista en centros de salud con tareas de atención al paciente, turnos y caja. Sin formación contable/administrativa formal ni conocimientos de inglés.",
                "puntos_fuertes": [
                    "Atención al público y gestión de turnos",
                    "Manejo básico de caja y recepción de insumos"
                ],
                "requisitos_faltantes": [
                    "Sin formación en Ciencias Económicas ni secundario comercial",
                    "Sin experiencia en compras corporativas, proveedores o cobranzas",
                    "Sin conocimientos de inglés"
                ],
                "archivo": "CV FÁTIMA SÁNCHEZ.pdf"
            },
            {
                "nombre_candidato": "Matías Adragna",
                "puntaje_compatibilidad": 25,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Analista de Soporte IT y Técnico con más de 15 años en Help Desk, hardware y redes. No se ajusta al área administrativa, contable ni de compras.",
                "puntos_fuertes": [
                    "Amplio dominio técnico de sistemas, soporte y seguridad de datos",
                    "Uso básico de SAP y paquete Office",
                    "Capacidad de resolución de incidencias bajo presión"
                ],
                "requisitos_faltantes": [
                    "Formación técnica/periodística, sin antecedentes en Ciencias Económicas",
                    "Sin experiencia en compras a proveedores, cobranzas o facturación contable",
                    "Nivel de inglés básico"
                ],
                "archivo": "CV_Matias_Adragna_Soporte_IT_Tecnico_Administrativo.pdf"
            },
            {
                "nombre_candidato": "Lucía Pereyra",
                "puntaje_compatibilidad": 22,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Desarrolladora de software frontend especializada en React, TypeScript y Next.js. Perfil orientado exclusivamente a programación web y diseño interactivo.",
                "puntos_fuertes": [
                    "Lógica algorítmica y dominio informático avanzado",
                    "Inglés técnico profesional B2",
                    "Tecnicatura Universitaria en Programación (UNSAM)"
                ],
                "requisitos_faltantes": ["Sin formación administrativa, económica ni comercial."],
                "archivo": "cv_06_frontend_developer.pdf"
            },
            {
                "nombre_candidato": "Lucía Fernández",
                "puntaje_compatibilidad": 20,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Diseñadora gráfica enfocada en identidad de marca y redes sociales. No cuenta con antecedentes en gestión administrativa ni el nivel de inglés requerido.",
                "puntos_fuertes": ["Manejo de herramientas de diseño visual y comunicación corporativa"],
                "requisitos_faltantes": ["No cumple los requisitos formativos, operativos ni lingüísticos solicitados."],
                "archivo": "CV_Lucia_Fernandez_Diseno.pdf"
            },
            {
                "nombre_candidato": "Nicolás Domínguez",
                "puntaje_compatibilidad": 20,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Diseñador de producto digital (UX/UI) y Design Systems en Figma. Trayectoria en investigación de usuarios y diseño de interfaces web/mobile.",
                "puntos_fuertes": [
                    "Título universitario en Diseño (FADU - UBA)",
                    "Nivel de inglés profesional C1",
                    "Capacidad de resolución de problemas orientada al usuario"
                ],
                "requisitos_faltantes": ["Perfil tecnológico y de diseño visual; sin antecedentes administrativos ni contables."],
                "archivo": "cv_07_ux_ui_designer.pdf"
            },
            {
                "nombre_candidato": "Carlos Alberto Benítez",
                "puntaje_compatibilidad": 18,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Operario calificado de depósito con carnet habilitante de autoelevadores (Res. SRT 960/15) y picking por radiofrecuencia. Experiencia 100% logística de planta.",
                "puntos_fuertes": [
                    "Manejo seguro de autoelevadores con certificación oficial Res. SRT 960/15",
                    "Experiencia en recepción, control de remitos de carga y estiba en altura",
                    "Disponibilidad horaria total y licencias de conducir al día"
                ],
                "requisitos_faltantes": [
                    "Sin formación administrativa ni estudios contables",
                    "Sin experiencia en compras de oficina, conciliaciones o facturación",
                    "Sin manejo de idioma inglés"
                ],
                "archivo": "CV_Carlos_Benitez_Operario.pdf"
            },
            {
                "nombre_candidato": "Tomás Benítez",
                "puntaje_compatibilidad": 18,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Diseñador Gráfico egresado de la UBA especializado en branding, identidad visual, preprensa y diseño editorial en Adobe Illustrator y Photoshop.",
                "puntos_fuertes": [
                    "Graduado de honor de la Universidad de Buenos Aires",
                    "Criterio estético y dominio avanzado de software de diseño",
                    "Inglés profesional B2"
                ],
                "requisitos_faltantes": ["Sin afinidad con funciones de administración, contabilidad o compras de oficina."],
                "archivo": "cv_20_disenador_grafico.pdf"
            },
            {
                "nombre_candidato": "Joaquín Vega",
                "puntaje_compatibilidad": 18,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Ingeniero en Informática (UBA) con especialización en QA Automation, desarrollo de scripts de prueba en Playwright/TypeScript e integración continua.",
                "puntos_fuertes": [
                    "Ingeniero en Informática egresado de la FIUBA",
                    "Inglés profesional avanzado C1",
                    "Capacidad analítica estructurada para pruebas de software"
                ],
                "requisitos_faltantes": ["Perfil puramente técnico de ingeniería de software; no aplica para tareas administrativas."],
                "archivo": "cv_11_qa_automation.pdf"
            },
            {
                "nombre_candidato": "Damián Godoy",
                "puntaje_compatibilidad": 16,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Administrador de Sistemas (SysAdmin) Linux y Redes corporativas. Experiencia en virtualización VMware/Proxmox, servidores y monitoreo con Zabbix.",
                "puntos_fuertes": [
                    "Sólido conocimiento en administración de servidores, seguridad e infraestructura",
                    "Certificación Red Hat (RHCSA) y Cisco CCNA",
                    "Inglés técnico avanzado B2"
                ],
                "requisitos_faltantes": ["Perfil exclusivamente de infraestructura IT; sin relación con áreas contables ni de compras."],
                "archivo": "cv_29_sysadmin_linux.pdf"
            },
            {
                "nombre_candidato": "Jeremías Fermín Jerez",
                "puntaje_compatibilidad": 15,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Técnico electrónico y estudiante de Ingeniería Industrial. Experiencia orientada a tornería, metrología, depósito de repuestos y construcción.",
                "puntos_fuertes": ["Conocimientos técnicos de precisión, metrología y organización de depósito"],
                "requisitos_faltantes": [
                    "Sin formación en Ciencias Económicas ni administración contable",
                    "Sin experiencia en gestión de cobranzas, compras o facturación",
                    "Sin nivel de inglés requerido"
                ],
                "archivo": "CV - JEREMIAS JEREZ 2.pdf"
            },
            {
                "nombre_candidato": "Ignacio Castillo",
                "puntaje_compatibilidad": 15,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Senior Data Engineer y Licenciado en Computación (UBA). Especialista en Big Data, Apache Spark, Airflow y arquitecturas Data Lakehouse en AWS.",
                "puntos_fuertes": [
                    "Formación científica de grado en Ciencias de la Computación (UBA)",
                    "Inglés bilingüe profesional C1",
                    "Dominio de bases de datos masivas y programación avanzada"
                ],
                "requisitos_faltantes": ["Perfil sénior de ingeniería de datos en la nube sin relación con secretariado o administración."],
                "archivo": "cv_17_data_engineer.pdf"
            },
            {
                "nombre_candidato": "Lucas Santillán",
                "puntaje_compatibilidad": 14,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Desarrollador de aplicaciones móviles en Flutter y React Native para iOS y Android. Especialista en arquitecturas limpias y APIs bancarias.",
                "puntos_fuertes": [
                    "Desarrollo de software y lógica de programación",
                    "Inglés intermedio B2+",
                    "Tecnicatura en Desarrollo de Software"
                ],
                "requisitos_faltantes": ["Sin antecedentes ni competencias en gestión administrativa, compras o finanzas."],
                "archivo": "cv_25_mobile_developer.pdf"
            },
            {
                "nombre_candidato": "Lic. Paula Medina",
                "puntaje_compatibilidad": 12,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Licenciada en Enfermería con Diploma de Honor (UBA) especializada en Cuidados Críticos (UTI), soporte vital avanzado y coordinación asistencial.",
                "puntos_fuertes": [
                    "Grado universitario con Diploma de Honor de la UBA",
                    "Liderazgo de equipos y gestión de historias clínicas hospitalarias",
                    "Trabajo bajo estricta presión y protocolos de bioseguridad"
                ],
                "requisitos_faltantes": [
                    "Profesional de la salud clínica; sin formación económica o administrativa de empresas",
                    "Sin experiencia en compras comerciales, cobranzas ni facturación contable"
                ],
                "archivo": "cv_14_enfermeria_salud.pdf"
            },
            {
                "nombre_candidato": "Fernando Ariel Gill Alfonso",
                "puntaje_compatibilidad": 10,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Técnico electromecánico con experiencia en matricería, moldes y mantenimiento de planta. No presenta afinidad con roles de oficina.",
                "puntos_fuertes": ["Sólido dominio de ajuste mecánico de precisión y control dimensional"],
                "requisitos_faltantes": [
                    "Perfil exclusivamente industrial/metalúrgico",
                    "Sin competencias ni formación administrativa o contable"
                ],
                "archivo": "CV Fernando Alfonzo.pdf"
            },
            {
                "nombre_candidato": "Dra. Brenda Peralta",
                "puntaje_compatibilidad": 10,
                "cumple_excluyentes": False,
                "veredicto": "Descartar",
                "resumen_perfil": "Bioquímica Clínica graduada con honores de la UBA, especialista en análisis clínicos automatizados, hematología, cultivos y normas ISO 15189.",
                "puntos_fuertes": [
                    "Título universitario de grado de la UBA con honores y matrícula nacional",
                    "Rigurosidad analítica y gestión de calidad bajo normas ISO 15189",
                    "Inglés técnico avanzado B2"
                ],
                "requisitos_faltantes": [
                    "Orientación puramente bioquímica y biomédica",
                    "Sin experiencia en funciones administrativas, comerciales o contables"
                ],
                "archivo": "cv_24_bioquimico_laboratorio.pdf"
            }
        ]

        total = len(resultados_precargados)
        progreso = st.progress(0, text="Iniciando evaluación de candidatos...")

        for i, item in enumerate(resultados_precargados):
            progreso.progress((i + 1) / total, text=f"Analizando ({i + 1}/{total}): {item['archivo']}")
            time.sleep(0.35)  # 0.35s por archivo (~19s total para los 55 CVs)

        progreso.empty()
        resultados = resultados_precargados

    # RAMA 2: PROCESAMIENTO REAL CON API (Cadena de respaldo secuencial y estable)
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