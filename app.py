from flask import Flask, render_template, render_template as flask_render_template, request, redirect, url_for, flash, session, send_file, jsonify, abort
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import inspect as sqlalchemy_inspect
from sqlalchemy.orm import deferred
from datetime import date, datetime, timedelta
from pathlib import Path
from functools import wraps
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import hashlib
import hmac
import io
import json
import math
import time
import os
import csv
import unicodedata
import re
import cuestionario as cq
from priorizacion import generar_matriz_priorizacion

BASE_DIR = Path(__file__).resolve().parent
app = Flask(__name__)

# Selección segura de la base de datos:
# 1) Si Render/entorno define DATABASE_URL, se usa esa base de datos administrada.
# 2) Si define DB_PATH, se usa esa ruta (por ejemplo, /var/data/consejo.cda.db
#    si se configura un disco persistente en Render).
# 3) En ejecución local, se prioriza la base histórica consejo.cda.db.
# 4) Si no existe la histórica pero sí la actual, se conserva diagnostico_archivistico.db.
# No se copian, borran ni sobrescriben bases de datos automáticamente.
_database_url = os.environ.get("DATABASE_URL", "").strip()
_db_path_env = os.environ.get("DB_PATH", "").strip()

if _database_url:
    _database_uri = _database_url.replace("postgres://", "postgresql://", 1)
else:
    if _db_path_env:
        _db_path = Path(_db_path_env).expanduser()
        if not _db_path.is_absolute():
            _db_path = BASE_DIR / _db_path
    else:
        _legacy_db = BASE_DIR / "consejo.cda.db"
        _current_db = BASE_DIR / "diagnostico_archivistico.db"
        if _legacy_db.exists():
            _db_path = _legacy_db
        elif _current_db.exists():
            _db_path = _current_db
        else:
            _db_path = _legacy_db
    _database_uri = "sqlite:///" + _db_path.as_posix()

app.config.update(
    SECRET_KEY=os.environ.get("SECRET_KEY", "solo-desarrollo-cambiar-SECRET_KEY"),
    SQLALCHEMY_DATABASE_URI=_database_uri,
    SQLALCHEMY_TRACK_MODIFICATIONS=False,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("RENDER", "").lower() == "true",
    # Límite de carga: PDF de actas y convalidaciones de hasta 10 MB.
    MAX_CONTENT_LENGTH=12 * 1024 * 1024,
    # Recupera conexiones inactivas o cerradas por Neon antes de usarlas.
    SQLALCHEMY_ENGINE_OPTIONS={
        "pool_pre_ping": True,
        "pool_recycle": 300,
    },
)

if os.environ.get("RENDER", "").lower() == "true" and not _database_url and not _db_path_env:
    app.logger.warning(
        "La aplicación usa SQLite dentro de la carpeta del proyecto en Render. "
        "Ese almacenamiento puede perderse al desplegar o reiniciar. Configure "
        "DATABASE_URL o DB_PATH apuntando a almacenamiento persistente."
    )

if os.environ.get("RENDER", "").lower() == "true" and not os.environ.get("SECRET_KEY"):
    app.logger.error(
        "SECRET_KEY no está configurada en Render: las sesiones usan una clave de desarrollo. "
        "Defina SECRET_KEY en Environment con un valor aleatorio largo."
    )

app.logger.info("Almacén de datos configurado: %s", "DATABASE_URL" if _database_url else _db_path)
db = SQLAlchemy(app)

# En Render configura SECRET_KEY y ADMIN_PASSWORD en Environment.
# Para conservar datos en producción, configura DATABASE_URL o DB_PATH en un disco persistente.

class Censo(db.Model):
    # codigo_entidad vincula cada registro al directorio oficial; respuestas_json guarda
    # las 10 preguntas nuevas y sus respuestas condicionales sin perder información.
    id=db.Column(db.Integer,primary_key=True); fecha=db.Column(db.DateTime,default=datetime.now,nullable=False)
    nombre_entidad=db.Column(db.String(300),nullable=False); naturaleza=db.Column(db.String(200),nullable=False); sector=db.Column(db.String(150),nullable=False); responsable=db.Column(db.String(300),nullable=False)
    q5=db.Column(db.String(200)); q6_json=db.Column(db.Text); q7=db.Column(db.String(250)); q8=db.Column(db.String(200)); q9=db.Column(db.String(250)); q10=db.Column(db.String(250))
    q11=db.Column(db.Text); q12=db.Column(db.Text); q13=db.Column(db.Text); q14=db.Column(db.String(250)); q15=db.Column(db.String(100)); q16=db.Column(db.String(200)); q17=db.Column(db.Text); q18=db.Column(db.String(100)); q19=db.Column(db.String(200)); q20=db.Column(db.Text); q21=db.Column(db.String(100)); respuestas_json=db.Column(db.Text); codigo_entidad=db.Column(db.String(50))
    puntaje_obtenido=db.Column(db.Float,default=0); variables_aplicables=db.Column(db.Float,default=0); ima=db.Column(db.Float,default=0); nivel=db.Column(db.String(100))

class BorradorDiagnostico(db.Model):
    """Almacena respuestas parciales sin alterar diagnósticos finalizados."""
    __tablename__ = 'borrador_diagnostico'
    id = db.Column(db.String(80), primary_key=True)
    datos_json = db.Column(db.Text, nullable=False, default='{}')
    estado = db.Column(db.String(20), nullable=False, default='borrador')
    creado_en = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    actualizado_en = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)
    censo_id = db.Column(db.Integer, nullable=True)
    # Entidad dueña del borrador cuando el diagnóstico exige ingreso con código.
    entidad_clave = db.Column(db.String(120), nullable=True, index=True)


class UbicacionEntidad(db.Model):
    """Caché persistente de ubicaciones resueltas bajo demanda desde una dirección.

    Es una tabla independiente; no modifica ni elimina registros de Censo.
    """
    __tablename__ = 'ubicacion_entidad'
    codigo_entidad = db.Column(db.String(50), primary_key=True)
    latitud = db.Column(db.Float, nullable=False)
    longitud = db.Column(db.Float, nullable=False)
    direccion_geocodificada = db.Column(db.Text)
    fuente = db.Column(db.String(100), default='Nominatim / OpenStreetMap')
    actualizado_en = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


INSTRUMENTOS=["CCD","TRD","TVD","PGD","PINAR","Inventarios documentales","Banco terminológico","Modelo de requisitos para documentos electrónicos","Sistema Integrado de Conservación (SIC)","Plan de Conservación Documental","Plan de Preservación Digital a Largo Plazo"]
INSTRUMENTO_OPCIONES=[("0","No cuenta"),("0.5","En elaboración, actualización o implementación parcial"),("1","Vigente / implementado"),("NA","N/A")]

# Rangos del mapa de calor del nivel de madurez (porcentaje).
# Bajo: 0 a 40 · Medio: más de 40 hasta 75 · Alto: más de 75 hasta 100.
UMBRAL_BAJO = 40
UMBRAL_MEDIO = 75

# Ordenada por jerarquía normativa (nivel 1 = leyes) y, dentro de cada nivel, por año.
NIVELES_NORMATIVOS = {
    1: "Leyes",
    2: "Decretos nacionales",
    3: "Acuerdos y resoluciones de autoridades nacionales",
    4: "Normas distritales",
    5: "Circulares y directrices",
    6: "Guías, instrumentos y referentes técnicos",
}

NORMATIVIDAD=[
    {
        "categoria": "Ley",
        "titulo": "Ley 594 de 2000",
        "descripcion": "Ley General de Archivos y marco general de la función archivística del Estado.",
        "estado": "Vigente",
        "icono": "bi-journal-text",
        "web": "https://normativa.archivogeneral.gov.co/ley-594-de-2000/",
        "pdf": "https://normativa.archivogeneral.gov.co/ley-594-de-2000/?pdf=41",
        "nivel": 1
    },
    {
        "categoria": "Ley",
        "titulo": "Ley 1581 de 2012",
        "descripcion": "Ley Estatutaria de protección de datos personales, aplicable a los documentos y bases de datos que gestionan las entidades.",
        "estado": "Vigente",
        "icono": "bi-person-lock",
        "web": "https://www.funcionpublica.gov.co/eva/gestornormativo/norma.php?i=49981",
        "pdf": "",
        "nivel": 1
    },
    {
        "categoria": "Ley",
        "titulo": "Ley 1712 de 2014",
        "descripcion": "Transparencia y derecho de acceso a la información pública nacional.",
        "estado": "Vigente",
        "icono": "bi-shield-check",
        "web": "https://normativa.archivogeneral.gov.co/ley-1712-de-2014/",
        "pdf": "https://normativa.archivogeneral.gov.co/ley-1712-de-2014/?pdf=418",
        "nivel": 1
    },
    {
        "categoria": "Decreto",
        "titulo": "Decreto 1080 de 2015",
        "descripcion": "Decreto Único Reglamentario del Sector Cultura, con disposiciones aplicables a la gestión documental.",
        "estado": "Vigente",
        "icono": "bi-file-earmark-ruled",
        "web": "https://normativa.archivogeneral.gov.co/decreto-1080-de-2015/",
        "pdf": "https://normativa.archivogeneral.gov.co/decreto-1080-de-2015/?pdf=33",
        "nivel": 2
    },
    {
        "categoria": "Resolución",
        "titulo": "Resolución MinTIC 1519 de 2020",
        "descripcion": "Lineamientos de publicación, accesibilidad web, transparencia, seguridad digital y datos abiertos.",
        "estado": "Fuente normativa oficial",
        "icono": "bi-gear",
        "web": "https://normograma.mintic.gov.co/mintic/compilacion/docs/resolucion_mintic_1519_2020.htm",
        "pdf": "",
        "nivel": 3
    },
    {
        "categoria": "Acuerdo",
        "titulo": "Acuerdo No. 001 de 2024",
        "descripcion": "Acuerdo Único de la Función Archivística y criterios técnicos y jurídicos para su implementación.",
        "estado": "Vigente",
        "icono": "bi-file-earmark-medical",
        "web": "https://normativa.archivogeneral.gov.co/acuerdo-no-001-del-2024/",
        "pdf": "https://normativa.archivogeneral.gov.co/wp-content/uploads/2024/04/2024-02_29_AcuerdoAGN-FIRMADO.pdf",
        "nivel": 3
    },
    {
        "categoria": "Decreto distrital",
        "titulo": "Decreto Distrital 0673 de 2024",
        "descripcion": "Crea y reglamenta el Consejo Distrital de Archivos de Medellín: integración, funciones, Presidencia, Secretaría Técnica, sesiones, quórum y actas.",
        "estado": "Vigente · texto oficial en la Gaceta Oficial del Distrito",
        "icono": "bi-buildings",
        "web": "",
        "pdf": "",
        "nivel": 4
    },
    {
        "categoria": "Circular",
        "titulo": "Circulares y directrices archivísticas del AGN",
        "descripcion": "Consulta las circulares y orientaciones oficiales emitidas para la función archivística.",
        "estado": "Repositorio oficial de circulares",
        "icono": "bi-file-earmark",
        "web": "https://normativa.archivogeneral.gov.co/inicio/circulares/",
        "pdf": "",
        "nivel": 5
    },
    {
        "categoria": "Guía y manual",
        "titulo": "Guías y manuales de gestión documental",
        "descripcion": "Material técnico de orientación para organizar, conservar y administrar documentos y archivos.",
        "estado": "Repositorio oficial de publicaciones",
        "icono": "bi-book",
        "web": "https://www.archivogeneral.gov.co/consulte/recursos/publicaciones",
        "pdf": "",
        "nivel": 6
    },
    {
        "categoria": "Instrumento archivístico",
        "titulo": "Instrumentos archivísticos: TRD, TVD, PGD y PINAR",
        "descripcion": "Referentes y recursos para la elaboración e implementación de instrumentos archivísticos.",
        "estado": "Portal oficial de instrumentos",
        "icono": "bi-stack",
        "web": "https://www.archivogeneral.gov.co/instrumentos-archivisticos",
        "pdf": "",
        "nivel": 6
    },
    {
        "categoria": "Referente técnico",
        "titulo": "MGDA – Modelo de Gestión Documental y Administración de Archivos",
        "descripcion": "Referente técnico del AGN para diagnóstico, gestión y desarrollo de la función archivística. La copia oficial disponible corresponde a la versión 2.0 (2020).",
        "estado": "Documento técnico",
        "icono": "bi-diagram-3",
        "web": "https://www.archivogeneral.gov.co/",
        "pdf": "https://www.archivogeneral.gov.co/sites/default/files/Estructura_Web/5_Consulte/Recursos/Publicacionees/V3_MGDA-min.pdf",
        "nivel": 6
    },
    {
        "categoria": "Documento formativo",
        "titulo": "Material formativo en gestión documental y archivos",
        "descripcion": "Recursos de capacitación, cartillas, infografías y materiales pedagógicos para fortalecer conocimientos archivísticos.",
        "estado": "Repositorio oficial de publicaciones",
        "icono": "bi-mortarboard",
        "web": "https://www.archivogeneral.gov.co/consulte/recursos/publicaciones",
        "pdf": "",
        "nivel": 6
    }
]

def calcular(form):
    total=aplicables=0.0; q6={}
    def add(v,m):
        nonlocal total,aplicables
        if v in m and m[v] is not None: total+=m[v]; aplicables+=1
    add(form.get('q5'),{'No (0)':0,'En proceso / Parcialmente (0,5)':.5,'Sí (1)':1})
    for i,n in enumerate(INSTRUMENTOS):
        v=form.get(f'q6_{i}',''); q6[n]=v; add(v,{'0':0,'0.5':.5,'1':1,'NA':None})
    add(form.get('q7'),{'No existen (0)':0,'En elaboración, actualización, trámite de convalidación o requieren actualización (0,5)':.5,'Convalidadas y vigentes (1)':1})
    add(form.get('q8'),{'No / No se ha verificado (0)':0,'Parcialmente (0,5)':.5,'Sí (1)':1})
    add(form.get('q9'),{'No existe (0)':0,'En elaboración, actualización, implementación o aprobado pero desactualizado (0,5)':.5,'Aprobado, vigente e implementado (1)':1})
    add(form.get('q10'),{'No existe (0)':0,'En elaboración, actualización, implementación o aprobado pero desactualizado (0,5)':.5,'Aprobado, vigente e implementado (1)':1})
    add(form.get('q14'),{'Con rezago significativo / No se ha diagnosticado (0)':0,'Parcialmente o mayormente organizados (0,5)':.5,'Organizados e inventariados (1)':1})
    add(form.get('q15'),{'No (0)':0,'Parcialmente (0,5)':.5,'Sí (1)':1,'No aplica (N/A)':None})
    add(form.get('q16'),{'Inadecuadas / No se ha evaluado (0)':0,'Parcialmente adecuadas (0,5)':.5,'Adecuadas (1)':1})
    add(form.get('q18'),{'No (0)':0,'Parcialmente (0,5)':.5,'Sí (1)':1,'No aplica (N/A)':None})
    add(form.get('q19'),{'No (0)':0,'En proceso / Implementación parcial (0,5)':.5,'Sí, implementado (1)':1})
    ima=round(total/aplicables*100,2) if aplicables else 0
    if ima <= UMBRAL_BAJO: nivel = 'Bajo'
    elif ima <= UMBRAL_MEDIO: nivel = 'Medio'
    else: nivel = 'Alto'
    return total,aplicables,ima,nivel,q6

@app.context_processor
def _umbrales():
    return {'UMBRAL_BAJO': UMBRAL_BAJO, 'UMBRAL_MEDIO': UMBRAL_MEDIO, 'cuestionario_version': cq.VERSION}


@app.route('/')
def index():
    # Cifras públicas de la portada: solo totales, nunca resultados de una entidad.
    try:
        diagnosticadas = db.session.query(db.func.count(db.distinct(db.func.coalesce(Censo.codigo_entidad, Censo.nombre_entidad)))).scalar() or 0
        cifras = {
            'entidades': len(obtener_directorio()),
            'diagnosticos': diagnosticadas,
            'convalidaciones': _convalidaciones_publicadas().count(),
            'actas': _actas_publicadas().count(),
        }
    except Exception:
        db.session.rollback()
        cifras = None
    conteos = {'actas': cifras['actas'] if cifras else 0,
               'convalidaciones': cifras['convalidaciones'] if cifras else 0,
               'documentos': 0}
    try:
        conteos['documentos'] = PublicacionCDA.query.filter(PublicacionCDA.publicado.is_(True),
                                                            PublicacionCDA.categoria != 'acta').count()
    except Exception:
        db.session.rollback()
    # La página de inicio es la del Consejo Distrital de Archivos (Sistema Distrital).
    return render_template('consejo.html', cifras=cifras, conteos=conteos)



# ============================================================
# DIRECTORIO DE ENTIDADES
# Búsqueda por nombre o sigla desde el archivo CSV
# ============================================================

def normalizar_texto(texto):
    """Normaliza texto para buscar sin distinguir mayúsculas ni tildes."""
    texto = str(texto or '').strip().lower()
    texto = unicodedata.normalize('NFD', texto)
    return ''.join(
        caracter for caracter in texto
        if unicodedata.category(caracter) != 'Mn'
    )


def clave_entidad(texto):
    """Clave comparable para vincular nombres del directorio y registros históricos."""
    texto = str(texto or '').split(' | NIT:', 1)[0]
    texto = normalizar_texto(texto)
    # Uniforma puntuación y espacios sin alterar el nombre que se muestra.
    texto = ''.join(caracter if caracter.isalnum() else ' ' for caracter in texto)
    return ' '.join(texto.split())


def _leer_directorio_csv():
    """Lee el directorio CSV ubicado junto a app.py."""
    ruta = BASE_DIR / 'directorio_entidades.csv'

    if not ruta.exists():
        app.logger.warning(
            "No se encontró el archivo del directorio: %s", ruta
        )
        return []

    try:
        with ruta.open(
            mode='r',
            encoding='utf-8-sig',
            newline=''
        ) as archivo:
            lector = csv.DictReader(archivo)

            if not lector.fieldnames:
                return []

            return [
                {
                    str(clave).strip(): str(valor or '').strip()
                    for clave, valor in fila.items()
                    if clave
                }
                for fila in lector
            ]

    except (OSError, csv.Error):
        app.logger.exception(
            "No fue posible leer el directorio de entidades."
        )
        return []


def obtener_directorio():
    """Directorio del CSV con el NIT que el administrador registre en la base de datos."""
    filas = _leer_directorio_csv()
    try:
        nits = {e.entidad_clave: e.nit for e in EntidadAcceso.query.filter(EntidadAcceso.nit.isnot(None)).all()}
    except Exception:
        db.session.rollback()
        nits = {}
    for fila in filas:
        fila.setdefault('NIT', '')
        nit = nits.get(clave_de_fila(fila))
        if nit:
            fila['NIT'] = nit
    return filas


def buscar_entidades_directorio(termino):
    """Busca coincidencias por nombre de entidad o sigla."""
    busqueda = normalizar_texto(termino)

    if len(busqueda) < 2:
        return []

    resultados = []

    for fila in obtener_directorio():
        campos = {
            normalizar_texto(clave): valor
            for clave, valor in fila.items()
        }

        nombre = next(
            (
                valor for clave, valor in campos.items()
                if 'entidad' in clave or 'nombre' in clave
            ),
            ''
        )

        sigla = next(
            (
                valor for clave, valor in campos.items()
                if 'sigla' in clave
            ),
            ''
        )

        if (
            busqueda in normalizar_texto(nombre)
            or busqueda in normalizar_texto(sigla)
        ):
            resultados.append({
                'nombre': nombre,
                'sigla': sigla,
                'datos': fila
            })

        if len(resultados) >= 10:
            break

    return resultados


CAMPOS_DIRECTORIO_PUBLICOS = {'Código', 'Entidad', 'Sigla', 'NIT', 'Grupo / relación con el Distrito',
                               'Clasificación preliminar', 'Ámbito', 'Responsable de gestión documental',
                               'Cargo Responsable GD', 'Correo Responsable GD', 'Teléfono Responsable GD',
                               'Sitio web / fuente', 'Dirección'}
# Campos del formulario que se precargan desde el directorio y la entidad puede corregir.
CAMPOS_CORREGIBLES = {'grupo_relacion': 'Grupo / relación', 'clasificacion_preliminar': 'Clasificación preliminar',
                      'ambito': 'Ámbito', 'responsable_nombre': 'Responsable', 'responsable_cargo': 'Cargo',
                      'responsable_correo': 'Correo', 'telefono_responsable': 'Teléfono',
                      'sitio_web': 'Sitio web', 'direccion_entidad': 'Dirección'}


@app.route('/api/buscar-entidades')
def api_buscar_entidades():
    """Devuelve coincidencias para el buscador del formulario."""
    termino = request.args.get('q', '').strip()

    if len(termino) < 2:
        return jsonify([])

    resultados = buscar_entidades_directorio(termino)

    # Solo se entregan las columnas del directorio, que son datos de contacto
    # institucional de servidores públicos (Ley 1712 de 2014, art. 9, lit. c).
    # La entidad puede corregirlos en el formulario.
    return jsonify([
        {
            'nombre': resultado['nombre'],
            'sigla': resultado['sigla'],
            'datos': {k: v for k, v in resultado['datos'].items() if k in CAMPOS_DIRECTORIO_PUBLICOS}
        }
        for resultado in resultados
    ])


# ============================================================
# AUTOGUARDADO Y RECUPERACIÓN DE BORRADORES
# La clave es un token aleatorio de alta entropía que actúa como código privado.
# Estas rutas solo crean/actualizan la tabla nueva; no alteran Censo ni ubicaciones.
# ============================================================


def validar_codigo_borrador(codigo):
    codigo = str(codigo or '').strip()
    return codigo if re.fullmatch(r'[A-Za-z0-9_-]{20,80}', codigo) else None


def _borrador_ajeno(borrador):
    """True si el borrador pertenece a otra entidad distinta de la que tiene la sesión."""
    if borrador is None or not borrador.entidad_clave or session.get('admin_authenticated'):
        return False
    entidad = session.get('entidad') or {}
    return entidad.get('clave') != borrador.entidad_clave


@app.route('/api/borrador/<codigo>', methods=['GET'])
def api_obtener_borrador(codigo):
    codigo = validar_codigo_borrador(codigo)
    if not codigo:
        return jsonify({'ok': False, 'mensaje': 'El código de recuperación no es válido.'}), 400
    try:
        borrador = db.session.get(BorradorDiagnostico, codigo)
        if borrador is None or _borrador_ajeno(borrador):
            return jsonify({'ok': True, 'existe': False, 'datos': {}})
        if borrador.estado != 'borrador':
            return jsonify({'ok': True, 'existe': True, 'finalizado': True, 'datos': {}})
        try:
            datos = json.loads(borrador.datos_json or '{}')
            if not isinstance(datos, dict):
                datos = {}
        except (TypeError, ValueError):
            datos = {}
        return jsonify({'ok': True, 'existe': True, 'finalizado': False,
                        'datos': datos, 'actualizado_en': borrador.actualizado_en.isoformat() if borrador.actualizado_en else None})
    except Exception:
        db.session.rollback()
        app.logger.exception('No se pudo recuperar el borrador')
        return jsonify({'ok': False, 'mensaje': 'No fue posible consultar el borrador en este momento.'}), 500


@app.route('/api/borrador/<codigo>', methods=['POST'])
def api_guardar_borrador(codigo):
    codigo = validar_codigo_borrador(codigo)
    if not codigo:
        return jsonify({'ok': False, 'mensaje': 'El código de recuperación no es válido.'}), 400
    payload = request.get_json(silent=True) or {}
    datos = payload.get('datos')
    if not isinstance(datos, dict):
        return jsonify({'ok': False, 'mensaje': 'El contenido del borrador no tiene un formato válido.'}), 400
    # No se guardan identificadores internos de control dentro de las respuestas.
    datos = {str(k)[:120]: v for k, v in datos.items() if k != 'borrador_id'}
    try:
        serializado = json.dumps(datos, ensure_ascii=False)
        if len(serializado.encode('utf-8')) > 500_000:
            return jsonify({'ok': False, 'mensaje': 'El borrador supera el tamaño permitido.'}), 413
        borrador = db.session.get(BorradorDiagnostico, codigo)
        if _borrador_ajeno(borrador):
            return jsonify({'ok': False, 'mensaje': 'Este borrador pertenece a otra entidad.'}), 403
        if borrador is None:
            borrador = BorradorDiagnostico(id=codigo, datos_json=serializado, estado='borrador',
                                           entidad_clave=(session.get('entidad') or {}).get('clave'))
            db.session.add(borrador)
        elif borrador.estado != 'borrador':
            return jsonify({'ok': False, 'finalizado': True,
                            'mensaje': 'Este diagnóstico ya fue finalizado y no admite más cambios.'}), 409
        else:
            borrador.datos_json = serializado
            borrador.actualizado_en = datetime.utcnow()
        db.session.commit()
        return jsonify({'ok': True, 'guardado': True,
                        'actualizado_en': borrador.actualizado_en.isoformat() if borrador.actualizado_en else None})
    except Exception:
        db.session.rollback()
        app.logger.exception('No se pudo guardar el borrador')
        return jsonify({'ok': False, 'mensaje': 'No se pudo guardar en la base de datos. Se conservará una copia local si hay espacio disponible.'}), 500


def _contexto_diagnostico(entidad):
    """Con ingreso por código: datos del directorio de la entidad y su borrador en curso."""
    if not entidad:
        return {}
    borrador = (BorradorDiagnostico.query.filter_by(entidad_clave=entidad['clave'], estado='borrador')
                .order_by(BorradorDiagnostico.actualizado_en.desc()).first())
    if borrador is None:
        borrador = BorradorDiagnostico(id=secrets.token_urlsafe(24), datos_json='{}', estado='borrador',
                                       entidad_clave=entidad['clave'])
        db.session.add(borrador)
        db.session.commit()
    fila = fila_por_clave(entidad['clave']) or {'Entidad': entidad['nombre']}
    return {'entidad_directorio': {k: fila.get(k, '') for k in CAMPOS_DIRECTORIO_PUBLICOS}, 'borrador_entidad': borrador.id}


@app.route('/diagnostico', methods=['GET', 'POST'])
def censo():
    entidad = None
    if acceso_requerido():
        entidad = entidad_en_sesion()
        if not entidad:
            return redirect(url_for('acceso_entidad', next=url_for('censo')))

    def render_template(plantilla, **kw):  # noqa: F811 - añade el contexto de la entidad
        return flask_render_template(plantilla, **_contexto_diagnostico(entidad), cuestionario=cq, **kw)

    if request.method == 'POST':
        obligatorios = ['nombre_entidad', 'nit_entidad', 'naturaleza', 'sector', 'direccion_entidad',
                        'responsable_nombre', 'responsable_cargo', 'responsable_correo']
        if any(not request.form.get(k, '').strip() for k in obligatorios):
            flash('Complete los campos obligatorios de la caracterización (C1 a C8).', 'danger')
            return render_template('diagnostico.html')
        sin_responder = [str(p['n']) for p in cq.PREGUNTAS
                         if request.form.get(p['campo'], '') not in dict(p['opciones'])]
        if sin_responder:
            flash(f'Faltan respuestas en las preguntas: {", ".join(sin_responder)}.', 'danger')
            return render_template('diagnostico.html')
        sin_justificar = [str(p['n']) for p in cq.PREGUNTAS for h in p['hijas']
                          if h.get('requerido') and cq.hija_visible(h, request.form, p)
                          and not request.form.get(h['campo'], '').strip()]
        if sin_justificar:
            flash(f'Falta la justificación en las preguntas: {", ".join(sin_justificar)}.', 'danger')
            return render_template('diagnostico.html')
        if not request.form.get('autoriza_datos'):
            flash('Debe autorizar el tratamiento de los datos personales para guardar el diagnóstico.', 'danger')
            return render_template('diagnostico.html')
        correo = request.form.get('responsable_correo', '').strip()
        if '@' not in correo or '.' not in correo.rsplit('@', 1)[-1]:
            flash('Ingrese un correo institucional válido.', 'danger')
            return render_template('diagnostico.html')

        # Vincula el registro al código del directorio solo si corresponde al nombre
        # seleccionado. Si se escribió el nombre manualmente, busca una coincidencia
        # exacta en el directorio; nunca conserva un código obsoleto de otra entidad.
        nombre = request.form.get('nombre_entidad', '').strip()
        codigo_enviado = request.form.get('codigo_entidad', '').strip()
        if entidad:
            # Con ingreso por código, la entidad la define la sesión y no el formulario.
            fila_sesion = fila_por_clave(entidad['clave']) or {}
            nombre = fila_sesion.get('Entidad') or entidad['nombre']
            codigo_enviado = str(fila_sesion.get('Código', '')).strip()
        nit_entidad = request.form.get('nit_entidad', '').strip()
        if entidad and (fila_por_clave(entidad['clave']) or {}).get('NIT'):
            nit_entidad = fila_por_clave(entidad['clave'])['NIT']
        filas_directorio = obtener_directorio()
        fila_codigo = next((fila for fila in filas_directorio
                            if str(fila.get('Código', '')).strip() == codigo_enviado), None)
        fila_nombre = next((fila for fila in filas_directorio
                            if clave_entidad(fila.get('Entidad', '')) == clave_entidad(nombre)), None)
        if fila_codigo and clave_entidad(fila_codigo.get('Entidad', '')) == clave_entidad(nombre):
            fila_vinculo = fila_codigo
        else:
            fila_vinculo = fila_nombre
        codigo_entidad = str(fila_vinculo.get('Código', '')).strip() if fila_vinculo else ''

        # Todas las respuestas del cuestionario (principales, hijas y caracterización) se
        # conservan en JSON. Las hijas que quedaron ocultas no se envían (campos deshabilitados).
        excluir = {'borrador_id', 'autoriza_datos', 'campos_corregidos'}
        respuestas = {k: (request.form.getlist(k) if len(request.form.getlist(k)) > 1 else request.form.get(k, ''))
                      for k in request.form.keys() if k not in excluir}
        for p in cq.PREGUNTAS:  # las selecciones múltiples se guardan siempre como lista
            for h in p['hijas']:
                if h['tipo'] == 'multi' and h['campo'] in respuestas and not isinstance(respuestas[h['campo']], list):
                    respuestas[h['campo']] = [respuestas[h['campo']]] if respuestas[h['campo']] else []
        sector = request.form.get('sector', '')
        if sector == 'Otro' and request.form.get('sector_otro', '').strip():
            sector = f"Otro: {request.form.get('sector_otro', '').strip()}"
        respuestas.update({
            'version': cq.VERSION,
            'codigo_entidad': codigo_entidad,
            'correo_acceso': entidad['correo'] if entidad else '',
            'nit_entidad': nit_entidad,
            'nombre_entidad': nombre,
            'correcciones_directorio': ', '.join(CAMPOS_CORREGIBLES[c] for c in request.form.get('campos_corregidos', '').split(',')
                                                 if c in CAMPOS_CORREGIBLES),
            'responsable_correo': correo,
        })
        total, aplicables, ima, nivel = calcular_diagnostico_nuevo(respuestas)
        nombre_y_nit = f"{nombre} | NIT: {nit_entidad}"
        responsable = ' | '.join([request.form.get('responsable_nombre','').strip(),
                                  request.form.get('responsable_cargo','').strip(), correo])
        r = Censo(
            nombre_entidad=nombre_y_nit,
            codigo_entidad=codigo_entidad or None,
            naturaleza=request.form.get('naturaleza','')[:200],
            sector=sector[:150], responsable=responsable[:300],
            respuestas_json=json.dumps(respuestas, ensure_ascii=False),
            puntaje_obtenido=total, variables_aplicables=aplicables, ima=ima, nivel=nivel)
        try:
            db.session.add(r)
            codigo_borrador = validar_codigo_borrador(request.form.get('borrador_id', ''))
            if codigo_borrador:
                borrador = db.session.get(BorradorDiagnostico, codigo_borrador)
                if borrador is not None and borrador.estado == 'borrador':
                    borrador.estado = 'finalizado'
                    borrador.censo_id = None  # se completa después del INSERT del registro
                    borrador.actualizado_en = datetime.utcnow()
            db.session.flush()
            if codigo_borrador:
                borrador = db.session.get(BorradorDiagnostico, codigo_borrador)
                if borrador is not None and borrador.estado == 'finalizado':
                    borrador.censo_id = r.id
            db.session.commit()
            session['mis_diagnosticos'] = (session.get('mis_diagnosticos', []) + [r.id])[-10:]
            flash('Diagnóstico guardado correctamente.', 'success')
            return redirect(url_for('resultado', censo_id=r.id))
        except Exception:
            db.session.rollback()
            app.logger.exception('No fue posible guardar el diagnóstico nuevo')
            flash('No fue posible guardar el diagnóstico. Revisa los registros de la aplicación.', 'danger')
            return render_template('diagnostico.html'), 500
    return render_template('diagnostico.html')


# Equivalencias de respuestas a puntaje por componente, en escala de tres valores:
# 0 = no cumple, 0,5 = avance parcial, 1 = cumple. «No se tiene información» y
# «No aplica» no se cuentan. Se comparten entre el cálculo del diagnóstico, la
# ficha de cada entidad y los indicadores del Sistema Distrital de Archivos.
MAPAS_COMPONENTES = {
        'estrategico': [
            ('q1_politica', {'No existe': 0, 'En desarrollo': 0.5, 'Aprobada': 0.5,
                             'En implementación': 0.5, 'Implementada y con seguimiento': 1,
                             'No se tiene información': None}),
            ('q2_seguimiento', {'No': 0, 'Parcialmente': 0.5, 'Sí': 1,
                                'No se tiene información': None}),
        ],
        'administracion_archivos': [
            ('q4_depositos', {'No': 0, 'Parcialmente': 0.5, 'Sí': 1,
                              'No se tiene información': None}),
        ],
        'procesos_documentales': [
            ('q5_trd', {'No elaboradas': 0, 'En elaboración': 0.5, 'En evaluación': 0.5,
                        'Convalidadas': 0.5, 'En implementación': 0.5, 'Implementadas': 1,
                        'Requieren actualización': 0.5, 'No aplica': None,
                        'No se tiene información': None}),
        ],
        'tecnologico': [
            ('q7_sgdea', {'No': 0, 'Parcialmente': 0.5, 'Sí': 1,
                          'No se tiene información': None}),
        ],
        'cultural': [
            ('q9_capacitacion', {'No se realiza': 0, 'Ocasional': 0.5,
                                 'Planificada y periódica': 1, 'No se tiene información': None}),
        ],
}


def puntajes_componentes(form):
    """Puntaje de 0 a 1 de cada componente (None si no hay respuestas puntuables).

    Los diagnósticos con el cuestionario V04 se calculan con cuestionario.py; los
    anteriores, con las equivalencias de MAPAS_COMPONENTES.
    """
    if cq.es_v04(form):
        return cq.puntajes_componentes(form)
    componentes = {}
    for componente, preguntas in MAPAS_COMPONENTES.items():
        valores = []
        for campo, opciones in preguntas:
            valor = opciones.get(form.get(campo, ''))
            if valor is not None:
                valores.append(valor)
        componentes[componente] = (sum(valores) / len(valores)) if valores else None
    return componentes


def calcular_diagnostico_nuevo(form):
    """Nivel de madurez: promedio de los cinco componentes (20 % cada uno).

    Cada componente es el promedio de sus preguntas (No = 0, Parcialmente = 0,5,
    Sí = 1); «N/A – No aplica» no se cuenta. Las preguntas de caracterización y las
    hijas son informativas y no suman puntaje.
    """
    componentes = puntajes_componentes(form)
    puntajes = [v for v in componentes.values() if v is not None]
    ima = round((sum(puntajes) / len(puntajes)) * 100, 2) if puntajes else 0
    total = round(sum(puntajes), 2)
    aplicables = float(len(puntajes))
    if ima <= UMBRAL_BAJO: nivel = 'Bajo'
    elif ima <= UMBRAL_MEDIO: nivel = 'Medio'
    else: nivel = 'Alto'
    return total, aplicables, ima, nivel

@app.route('/resultado/<int:censo_id>')
def resultado(censo_id):
    # Solo lo ve la persona que envió el diagnóstico (en su sesión) o el equipo administrador.
    registro = Censo.query.get_or_404(censo_id)
    if not _puede_ver_censo(registro):
        abort(404)
    respuestas, actual, componentes = _puntajes_de(registro)
    return render_template('resultado.html', censo=registro, formulario_actual=actual,
                           tabla_componentes=tabla_componentes(respuestas, componentes) if actual else [])

@app.route('/normatividad')
def normatividad(): return render_template('normatividad.html',normatividad=NORMATIVIDAD,niveles=NIVELES_NORMATIVOS)

@app.route('/ayuda')
def ayuda():
    return render_template('ayuda.html')


@app.route('/patrimonio-documental')
def patrimonio_documental():
    return render_template('patrimonio.html')


@app.route('/archivos-derechos-humanos')
def archivos_ddhh():
    return render_template('archivos_ddhh.html')


@app.route('/diagnostico/guia')
def guia_diagnostico():
    return render_template('guia_diagnostico.html', cuestionario=cq)



ADMIN_INACTIVIDAD_SEG = 60 * 60
_INTENTOS_LOGIN = {}  # IP -> lista de marcas de tiempo de intentos fallidos


def _ip_cliente():
    return (request.headers.get('X-Forwarded-For', request.remote_addr or '').split(',')[0]).strip()


def admin_required(view):
    """Exige una sesión administrativa válida para las rutas protegidas."""
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        if not session.get("admin_authenticated"):
            return redirect(url_for("admin_login", next=request.path))
        # Cierre automático tras 60 minutos sin actividad.
        if time.time() - session.get("admin_ultima_actividad", 0) > ADMIN_INACTIVIDAD_SEG:
            session.clear()
            flash("La sesión administrativa se cerró por inactividad.", "warning")
            return redirect(url_for("admin_login", next=request.path))
        session["admin_ultima_actividad"] = time.time()
        return view(*args, **kwargs)
    return wrapped_view


# MÓDULO NUEVO: mapa de entidades. No modifica las rutas existentes.
@app.route('/admin/mapa-entidades')
@admin_required
def mapa_entidades():
    # La clave es pública en el navegador: restríngela por dominio en Google Cloud.
    # No almacena coordenadas de Google de forma permanente; el mapa consulta la ubicación al seleccionar una entidad.
    return render_template(
        'mapa_entidades.html',
        google_maps_api_key=os.environ.get('GOOGLE_MAPS_API_KEY', '').strip()
    )


@app.route('/api/admin/mapa-entidades')
@admin_required
def api_mapa_entidades():
    """Entrega al mapa el directorio y el último diagnóstico por código de entidad.

    El código del directorio es el vínculo principal. Para registros históricos que
    no tengan código, se usa como respaldo el nombre normalizado de la entidad.
    Las coordenadas verificadas en el directorio tienen prioridad; las ubicaciones
    geocodificadas bajo demanda se conservan en una tabla independiente.
    """
    directorio = obtener_directorio()
    registros = Censo.query.order_by(Censo.fecha.desc(), Censo.id.desc()).all()

    clave_nombre = clave_entidad

    por_codigo = {}
    por_nombre = {}

    # Al venir en orden descendente, setdefault conserva el diagnóstico más reciente.
    for registro in registros:
        try:
            respuestas = json.loads(registro.respuestas_json or '{}')
            if not isinstance(respuestas, dict):
                respuestas = {}
        except (TypeError, ValueError):
            respuestas = {}

        codigo = str(registro.codigo_entidad or respuestas.get('codigo_entidad', '') or '').strip()
        nombre = clave_nombre(registro.nombre_entidad)
        diagnostico = {
            'id': registro.id,
            'ima': round(float(registro.ima or 0), 2),
            'nivel': registro.nivel or 'Sin clasificar',
            'fecha': registro.fecha.strftime('%Y-%m-%d') if registro.fecha else '',
            'nombre_registrado': registro.nombre_entidad or ''
        }
        if codigo:
            por_codigo.setdefault(codigo, diagnostico)
        if nombre:
            por_nombre.setdefault(nombre, diagnostico)

    def leer_coordenada(fila, posibles_nombres, minimo, maximo):
        raw = next((fila.get(nombre) for nombre in posibles_nombres if fila.get(nombre)), '')
        try:
            valor = float(str(raw).strip().replace(',', '.'))
            return valor if minimo <= valor <= maximo else None
        except (TypeError, ValueError):
            return None

    salida = []
    ubicaciones_guardadas = {
        ubicacion.codigo_entidad: ubicacion
        for ubicacion in UbicacionEntidad.query.all()
    }
    for fila in directorio:
        codigo = str(fila.get('Código', '') or '').strip()
        nombre = str(fila.get('Entidad', '') or '').strip()
        diagnostico = por_codigo.get(codigo) if codigo else None
        if diagnostico is None:
            diagnostico = por_nombre.get(clave_nombre(nombre))

        # Si el directorio ya tiene coordenadas, estas tienen prioridad. Si no,
        # se utilizan las coordenadas geocodificadas y guardadas previamente.
        latitud = leer_coordenada(fila, ['Latitud', 'latitud', 'Latitude'], -90, 90)
        longitud = leer_coordenada(fila, ['Longitud', 'longitud', 'Longitude'], -180, 180)
        fuente_ubicacion = 'directorio' if latitud is not None and longitud is not None else ''
        direccion_geocodificada = ''
        ubicacion_guardada = ubicaciones_guardadas.get(codigo) if codigo else None
        if (latitud is None or longitud is None) and ubicacion_guardada:
            latitud = ubicacion_guardada.latitud
            longitud = ubicacion_guardada.longitud
            fuente_ubicacion = ubicacion_guardada.fuente or 'Nominatim / OpenStreetMap'
            direccion_geocodificada = ubicacion_guardada.direccion_geocodificada or ''

        salida.append({
            'codigo': codigo,
            'nombre': nombre,
            'sigla': fila.get('Sigla', ''),
            'clasificacion': fila.get('Clasificación preliminar', 'Sin clasificación'),
            'ambito': fila.get('Ámbito', ''),
            'grupo': fila.get('Grupo / relación con el Distrito', ''),
            'direccion': fila.get('Dirección', ''),
            'responsable': fila.get('Responsable de gestión documental', ''),
            'cargo': fila.get('Cargo Responsable GD', ''),
            'correo': fila.get('Correo Responsable GD', ''),
            'telefono': fila.get('Teléfono Responsable GD', ''),
            'web': fila.get('Sitio web / fuente', ''),
            'latitud': latitud,
            'longitud': longitud,
            'ubicacion_fuente': fuente_ubicacion,
            'direccion_geocodificada': direccion_geocodificada,
            'diagnostico': diagnostico
        })

    con_diagnostico = sum(1 for entidad in salida if entidad['diagnostico'] is not None)
    return jsonify({
        'entidades': salida,
        'total': len(salida),
        'con_diagnostico': con_diagnostico,
        'sin_diagnostico': len(salida) - con_diagnostico
    })


@app.route('/api/admin/mapa-entidades/ubicacion', methods=['POST'])
@admin_required
def api_guardar_ubicacion_entidad():
    """Guarda en una tabla independiente la ubicación encontrada para una entidad.

    Solo acepta códigos presentes en directorio_entidades.csv y coordenadas válidas.
    No modifica el registro del diagnóstico ni el directorio original.
    """
    datos = request.get_json(silent=True) or {}
    codigo = str(datos.get('codigo', '') or '').strip()
    if not codigo:
        return jsonify({'error': 'Falta el código de la entidad.'}), 400

    fila = next((f for f in obtener_directorio()
                 if str(f.get('Código', '') or '').strip() == codigo), None)
    if not fila:
        return jsonify({'error': 'El código no existe en el directorio.'}), 404

    try:
        latitud = float(datos.get('latitud'))
        longitud = float(datos.get('longitud'))
    except (TypeError, ValueError):
        return jsonify({'error': 'Las coordenadas no son válidas.'}), 400

    if not (-90 <= latitud <= 90 and -180 <= longitud <= 180):
        return jsonify({'error': 'Las coordenadas están fuera de rango.'}), 400

    ubicacion = UbicacionEntidad.query.filter_by(codigo_entidad=codigo).first()
    if ubicacion is None:
        ubicacion = UbicacionEntidad(codigo_entidad=codigo)
    ubicacion.latitud = latitud
    ubicacion.longitud = longitud
    ubicacion.direccion_geocodificada = str(datos.get('direccion_geocodificada', '') or '')[:2000]
    ubicacion.fuente = 'Nominatim / OpenStreetMap'

    try:
        db.session.add(ubicacion)
        db.session.commit()
    except Exception:
        db.session.rollback()
        app.logger.exception('No se pudo guardar la ubicación de la entidad %s', codigo)
        return jsonify({'error': 'No se pudo guardar la ubicación.'}), 500

    return jsonify({
        'ok': True,
        'codigo': codigo,
        'latitud': ubicacion.latitud,
        'longitud': ubicacion.longitud,
        'fuente': ubicacion.fuente
    })


@app.route('/api/admin/mapa-entidades/ubicaciones', methods=['DELETE'])
@admin_required
def api_reiniciar_ubicaciones_entidades():
    """Borra solamente la caché de coordenadas geocodificadas.

    No elimina ni modifica diagnósticos, ni el CSV del directorio.
    Se usa para corregir ubicaciones guardadas incorrectamente.
    """
    try:
        cantidad = UbicacionEntidad.query.delete(synchronize_session=False)
        db.session.commit()
        return jsonify({'ok': True, 'eliminadas': cantidad})
    except Exception:
        db.session.rollback()
        app.logger.exception('No se pudo reiniciar la caché de ubicaciones')
        return jsonify({'error': 'No se pudo limpiar la caché de coordenadas.'}), 500


@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        configured_password = os.environ.get('ADMIN_PASSWORD', '')
        if not configured_password:
            flash('El acceso administrativo no está configurado. Contacte al administrador técnico.', 'danger')
            return render_template('admin_login.html'), 503

        ip = _ip_cliente()
        recientes = [t for t in _INTENTOS_LOGIN.get(ip, []) if time.time() - t < 900]
        _INTENTOS_LOGIN[ip] = recientes
        if len(recientes) >= 5:
            flash('Demasiados intentos fallidos. Espere 15 minutos e intente de nuevo.', 'danger')
            return render_template('admin_login.html'), 429

        if hmac.compare_digest(request.form.get('password', '').encode(), configured_password.encode()):
            _INTENTOS_LOGIN.pop(ip, None)
            session.clear()
            session['admin_authenticated'] = True
            session['admin_ultima_actividad'] = time.time()
            flash('Acceso administrativo autorizado.', 'success')
            destination = request.args.get('next', '')
            # Solo permitir rutas locales; impedir redirecciones externas.
            if not destination.startswith('/') or destination.startswith('//'):
                destination = url_for('admin')
            return redirect(destination)

        _INTENTOS_LOGIN[ip].append(time.time())
        app.logger.warning('Intento fallido de acceso administrativo desde %s', ip)
        flash('Contraseña incorrecta.', 'danger')

    return render_template('admin_login.html')


@app.route('/admin/logout', methods=['POST'])
def admin_logout():
    session.clear()
    flash('Sesión administrativa cerrada.', 'success')
    return redirect(url_for('admin_login'))


@app.route('/admin')
@admin_required
def admin():
    # Del más crítico al más avanzado: índice ascendente; a igual índice, el más reciente primero.
    registros=Censo.query.order_by(Censo.ima.asc(), Censo.fecha.desc(), Censo.id.desc()).all()
    promedio=round(sum(x.ima for x in registros)/len(registros),2) if registros else 0
    niveles={}
    sectores=sorted({(x.sector or 'Sin sector') for x in registros})
    for x in registros:
        niveles[x.nivel or 'Sin clasificar']=niveles.get(x.nivel or 'Sin clasificar',0)+1

    def semaforo(ima):
        # Clasificación simplificada en tres niveles: Bajo, Medio y Alto.
        if ima <= UMBRAL_BAJO: return {'label':'Bajo','color':'#dc3545','class':'low','texto':'#b42318'}
        if ima <= UMBRAL_MEDIO: return {'label':'Medio','color':'#e0a800','class':'medium','texto':'#7a5a00'}
        return {'label':'Alto','color':'#198754','class':'high','texto':'#14532d'}

    semaforos={
        'Bajo':sum(1 for x in registros if (x.ima or 0) <= UMBRAL_BAJO),
        'Medio':sum(1 for x in registros if UMBRAL_BAJO < (x.ima or 0) <= UMBRAL_MEDIO),
        'Alto':sum(1 for x in registros if (x.ima or 0) > UMBRAL_MEDIO),
    }
    ranking=[{'nombre':_sda_nombre(x),'ima':round(x.ima or 0,2),'nivel':semaforo(x.ima or 0)['label'],'semaforo':semaforo(x.ima or 0)['label']} for x in sorted(registros,key=lambda r:r.ima or 0,reverse=True)]

    # Avance promedio por componente, calculado con las respuestas del formulario vigente.
    # Si no hay respuestas nuevas, se usa el promedio de instrumentos (pregunta 6) de
    # los censos anteriores para no perder esos registros.
    sumas={clave:[0.0,0] for clave,_ in SDA_COMPONENTES}
    instrumentos_totales={nombre:[0.0,0] for nombre in INSTRUMENTOS}
    for x in registros:
        try: respuestas=json.loads(x.respuestas_json or '{}')
        except (TypeError,ValueError): respuestas={}
        if isinstance(respuestas, dict) and respuestas:
            for clave,valor in puntajes_componentes(respuestas).items():
                if valor is not None and clave in sumas:
                    sumas[clave][0]+=valor; sumas[clave][1]+=1
        try: antiguas=json.loads(x.q6_json or '{}')
        except (TypeError,ValueError): antiguas={}
        if isinstance(antiguas, dict):
            for nombre,valor in antiguas.items():
                if nombre in instrumentos_totales and valor in ('0','0.5','1'):
                    instrumentos_totales[nombre][0]+=float(valor); instrumentos_totales[nombre][1]+=1
    instrumentos_labels=[]; instrumentos_valores=[]
    for clave,nombre in SDA_COMPONENTES:
        suma,cantidad=sumas[clave]
        if cantidad:
            instrumentos_labels.append(nombre); instrumentos_valores.append(round(suma/cantidad*100,2))
    if not instrumentos_labels:
        for nombre,(suma,cantidad) in instrumentos_totales.items():
            if cantidad:
                instrumentos_labels.append(nombre); instrumentos_valores.append(round(suma/cantidad*100,2))
    # Evolución: entidades con dos o más diagnósticos, primer resultado frente al último.
    por_entidad = {}
    for x in sorted(registros, key=lambda r: (r.fecha or datetime.min, r.id)):
        por_entidad.setdefault(_ficha_clave(x), []).append(x)
    evolucion = []
    for lista in por_entidad.values():
        if len(lista) > 1:
            primero, ultimo = lista[0], lista[-1]
            evolucion.append({'nombre': _sda_nombre(ultimo), 'id': ultimo.id, 'n': len(lista),
                              'inicial': primero.ima or 0, 'actual': ultimo.ima or 0,
                              'cambio': round((ultimo.ima or 0) - (primero.ima or 0), 1)})
    evolucion.sort(key=lambda e: -e['cambio'])
    resumen_evolucion = {'mejoraron': sum(1 for e in evolucion if e['cambio'] > 0),
                         'iguales': sum(1 for e in evolucion if e['cambio'] == 0),
                         'bajaron': sum(1 for e in evolucion if e['cambio'] < 0),
                         'promedio': round(sum(e['cambio'] for e in evolucion) / len(evolucion), 1) if evolucion else 0}
    return render_template('admin.html',registros=registros,promedio=promedio,niveles=niveles,
        sectores=sectores,semaforos=semaforos,ranking=ranking,instrumentos_labels=instrumentos_labels,
        instrumentos_valores=instrumentos_valores,semaforo_fn=semaforo,evolucion=evolucion,
        resumen_evolucion=resumen_evolucion, clave_de=_ficha_clave,
        estados_acomp={k: _estado_acompanamiento(v) for k, v in _acompanamientos_por_entidad().items()})


@app.route('/admin/eliminar-diagnostico/<int:censo_id>', methods=['POST'])
@admin_required
def eliminar_diagnostico(censo_id):
    """Elimina solo el diagnóstico seleccionado, previa confirmación en el panel."""
    registro = db.session.get(Censo, censo_id)
    if registro is None:
        flash('El diagnóstico seleccionado ya no existe.', 'warning')
        return redirect(url_for('admin'))

    nombre = registro.nombre_entidad
    try:
        db.session.delete(registro)
        db.session.commit()
        flash(f'Se eliminó el diagnóstico de «{nombre}».', 'success')
    except Exception:
        db.session.rollback()
        app.logger.exception('No se pudo eliminar el diagnóstico %s', censo_id)
        flash('No fue posible eliminar el diagnóstico. No se realizaron cambios.', 'danger')
    return redirect(url_for('admin'))


@app.route('/admin/exportar-priorizacion')
@admin_required
def exportar_priorizacion():
    """Matriz de priorización: diagnóstico más reciente de cada entidad, por componente."""
    vistos, entidades = set(), []
    acomp = _acompanamientos_por_entidad()
    for r in Censo.query.order_by(Censo.fecha.desc(), Censo.id.desc()).all():
        clave = (r.codigo_entidad or '').strip() or clave_entidad(_sda_nombre(r))
        if clave in vistos:
            continue
        vistos.add(clave)
        respuestas, actual, componentes = _puntajes_de(r)
        if not actual:
            continue  # registros muy antiguos sin respuestas por componente
        conteos = {}
        if cq.es_v04(respuestas):
            for fila, (clave_c, _) in zip(tabla_componentes(respuestas, componentes), SDA_COMPONENTES):
                conteos[clave_c] = {'si': fila['si'], 'parcial': fila['parcial'], 'no': fila['no'], 'na': fila['na']}
        nit = respuestas.get('nit_entidad') or ''
        if not nit and '| NIT:' in (r.nombre_entidad or ''):
            nit = r.nombre_entidad.split('| NIT:')[-1].strip()
        entidades.append({
            'nombre': _sda_nombre(r), 'nit': nit, 'codigo': r.codigo_entidad or '',
            'sector': (r.sector or '').strip() or 'Sin sector',
            'naturaleza': (r.naturaleza or '').strip() or 'Sin dato',
            'fecha': r.fecha, 'version': cq.VERSION if cq.es_v04(respuestas) else 'Anterior',
            'ima': r.ima or 0, 'componentes': componentes, 'conteos': conteos,
            'acompanamiento': _estado_acompanamiento(acomp.get(clave)),
        })
    salida = generar_matriz_priorizacion(entidades, SDA_COMPONENTES, UMBRAL_BAJO, UMBRAL_MEDIO)
    nombre = f"Matriz_priorizacion_CDA_{datetime.now():%Y-%m-%d}.xlsx"
    return send_file(salida, as_attachment=True, download_name=nombre,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/admin/exportar-excel')
@admin_required
def exportar_excel():
    """Exporta los registros del diagnóstico a un libro Excel con dos hojas."""
    registros = Censo.query.order_by(Censo.fecha.desc()).all()
    wb = Workbook()
    ws = wb.active
    ws.title = 'Matriz general'

    encabezados = [
        'ID de registro', 'Fecha de registro', 'Entidad / NIT', 'Naturaleza',
        'Sector', 'Responsable', 'P5 - Instrumentos archivísticos',
    ]
    encabezados.extend([f'P6 - {nombre}' for nombre in INSTRUMENTOS])
    encabezados.extend([
        'P7 - TRD', 'P8 - Organización del archivo', 'P9 - PGD',
        'P10 - PINAR', 'P11', 'P12', 'P13', 'P14 - Organización e inventarios',
        'P15', 'P16 - Conservación', 'P17', 'P18', 'P19', 'P20', 'P21',
        'Puntaje obtenido', 'Variables aplicables', 'Nivel de madurez (%)', 'Clasificación',
    ])
    ws.append(encabezados)

    for r in registros:
        try:
            instrumentos = json.loads(r.q6_json or '{}')
            if not isinstance(instrumentos, dict):
                instrumentos = {}
        except (TypeError, ValueError):
            instrumentos = {}

        fila = [
            r.id,
            r.fecha.strftime('%Y-%m-%d %H:%M:%S') if r.fecha else '',
            r.nombre_entidad, r.naturaleza, r.sector, r.responsable, r.q5,
        ]
        fila.extend([instrumentos.get(nombre, '') for nombre in INSTRUMENTOS])
        fila.extend([
            r.q7, r.q8, r.q9, r.q10, r.q11, r.q12, r.q13, r.q14, r.q15,
            r.q16, r.q17, r.q18, r.q19, r.q20, r.q21,
            r.puntaje_obtenido, r.variables_aplicables, r.ima, r.nivel,
        ])
        ws.append(fila)

    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='17365D')
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 42
    for column_cells in ws.columns:
        max_length = min(max(len(str(cell.value or '')) for cell in column_cells) + 2, 42)
        ws.column_dimensions[get_column_letter(column_cells[0].column)].width = max(max_length, 14)

    resumen = wb.create_sheet('Resumen nivel de madurez')
    resumen.append([
        'ID de registro', 'Fecha de registro', 'Entidad / NIT', 'Naturaleza',
        'Sector', 'Puntaje obtenido', 'Variables aplicables', 'Nivel de madurez (%)', 'Clasificación'
    ])
    for r in registros:
        resumen.append([
            r.id,
            r.fecha.strftime('%Y-%m-%d %H:%M:%S') if r.fecha else '',
            r.nombre_entidad, r.naturaleza, r.sector,
            r.puntaje_obtenido, r.variables_aplicables, r.ima, r.nivel,
        ])
    for cell in resumen[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='17365D')
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    resumen.freeze_panes = 'A2'
    resumen.auto_filter.ref = resumen.dimensions
    for column_cells in resumen.columns:
        max_length = min(max(len(str(cell.value or '')) for cell in column_cells) + 2, 42)
        resumen.column_dimensions[get_column_letter(column_cells[0].column)].width = max(max_length, 14)

    # Hoja adicional con todas las respuestas del formulario nuevo, incluidas
    # las preguntas condicionales, evidencias y datos del directorio.
    respuestas_registros = []
    campos_respuesta = set()
    for registro in registros:
        try:
            respuestas = json.loads(registro.respuestas_json or '{}')
            if not isinstance(respuestas, dict):
                respuestas = {}
        except (TypeError, ValueError):
            respuestas = {}
        respuestas_registros.append((registro, respuestas))
        campos_respuesta.update(respuestas.keys())

    columnas_prioritarias = [
        'codigo_entidad', 'sigla_entidad', 'grupo_relacion',
        'clasificacion_preliminar', 'ambito', 'nit_entidad',
        'telefono_responsable', 'sitio_web', 'direccion_entidad', 'correcciones_directorio',
        'q1_politica', 'q1_acto', 'q1_acto_detalle', 'q1_etapa', 'q1_razon', 'evidencia_q1',
        'q2_seguimiento', 'q2_mecanismo', 'q2_pgd_formulado', 'evidencia_q2',
        'q3_personal', 'q3_observaciones', 'q4_depositos', 'q4_riesgos', 'evidencia_q4',
        'q5_trd', 'q5_avance', 'q5_evaluacion', 'q5_evidencia_estado', 'q5_actualizacion', 'evidencia_q5',
        'q6_procesos', 'evidencia_q6', 'q7_sgdea', 'q7_sistema_detalle', 'q7_soporte', 'q7_limitacion', 'evidencia_q7',
        'q8_protecciones', 'evidencia_q8', 'q9_capacitacion', 'q9_detalle', 'q9_barreras', 'evidencia_q9',
        'q10_prioridad', 'evidencia_q10',
    ]
    campos_ordenados = [campo for campo in columnas_prioritarias if campo in campos_respuesta]
    campos_ordenados.extend(sorted(campos_respuesta - set(campos_ordenados)))

    respuestas_ws = wb.create_sheet('Respuestas completas')
    respuestas_ws.append([
        'ID de registro', 'Fecha de registro', 'Código de entidad',
        'Entidad / NIT', 'Nivel de madurez (%)', 'Clasificación', *campos_ordenados
    ])
    for registro, respuestas in respuestas_registros:
        valores = []
        for campo in campos_ordenados:
            valor = respuestas.get(campo, '')
            if isinstance(valor, (list, tuple)):
                valor = '; '.join(str(elemento) for elemento in valor)
            elif isinstance(valor, dict):
                valor = json.dumps(valor, ensure_ascii=False)
            valores.append(valor)
        respuestas_ws.append([
            registro.id,
            registro.fecha.strftime('%Y-%m-%d %H:%M:%S') if registro.fecha else '',
            registro.codigo_entidad or respuestas.get('codigo_entidad', ''),
            registro.nombre_entidad,
            registro.ima,
            registro.nivel,
            *valores,
        ])
    for cell in respuestas_ws[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='17365D')
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    respuestas_ws.freeze_panes = 'A2'
    respuestas_ws.auto_filter.ref = respuestas_ws.dimensions
    respuestas_ws.row_dimensions[1].height = 48
    for column_cells in respuestas_ws.columns:
        max_length = min(max(len(str(cell.value or '')) for cell in column_cells) + 2, 42)
        respuestas_ws.column_dimensions[get_column_letter(column_cells[0].column)].width = max(max_length, 14)

    # Hoja principal con el cuestionario vigente: una fila por diagnóstico V04, con el
    # resultado por componente y todas las respuestas en el orden del formulario.
    v04 = wb.create_sheet(f'Diagnóstico {cq.VERSION}', 0)
    caracterizacion = [('naturaleza', 'C1. Naturaleza jurídica'), ('nombre_entidad', 'C2. Nombre de la entidad'),
                       ('nit_entidad', 'NIT'), ('sector', 'C3. Sector'), ('sector_otro', 'C3. Otro sector'),
                       ('direccion_entidad', 'C4. Dirección'), ('telefono_entidad', 'C5. Teléfono institucional'),
                       ('sitio_web', 'C6. Sitio web'), ('rep_nombre', 'C7. Representante legal'), ('rep_cargo', 'C7. Cargo'),
                       ('rep_correo', 'C7. Correo'), ('rep_telefono', 'C7. Teléfono'),
                       ('responsable_nombre', 'C8. Responsable de gestión documental'), ('responsable_cargo', 'C8. Cargo'),
                       ('responsable_dependencia', 'C8. Dependencia'), ('responsable_correo', 'C8. Correo'),
                       ('telefono_responsable', 'C8. Teléfono'), ('fecha_inicial', 'C9. Fecha inicial'),
                       ('fecha_final', 'C9. Fecha final'), ('fechas_desconocidas', 'C9. Fecha desconocida'),
                       ('fechas_observaciones', 'C9. Observaciones'), ('codigo_entidad', 'Código en el directorio'),
                       ('correcciones_directorio', 'Datos corregidos frente al directorio')]
    campos_v04 = cq.campos_formulario()
    v04.append(['ID de registro', 'Fecha de registro', *[e for _, e in caracterizacion],
                *[f"Componente {c['num']}. {c['nombre']} (%)" for c in cq.COMPONENTES],
                'Nivel de madurez (%)', 'Clasificación', *[e for _, e in campos_v04]])
    for registro, respuestas in respuestas_registros:
        if not cq.es_v04(respuestas):
            continue
        comps = cq.puntajes_componentes(respuestas)
        fila = [registro.id, registro.fecha.strftime('%Y-%m-%d %H:%M:%S') if registro.fecha else '']
        for campo, _ in caracterizacion:
            fila.append('Sí' if campo == 'fechas_desconocidas' and respuestas.get(campo) else respuestas.get(campo, ''))
        fila.extend(round(comps[c['clave']] * 100, 1) if comps.get(c['clave']) is not None else ''
                    for c in cq.COMPONENTES)
        fila.extend([registro.ima, registro.nivel])
        for campo, _ in campos_v04:
            valor = respuestas.get(campo, '')
            if campo in cq.POR_CAMPO:
                valor = cq.etiqueta_opcion(cq.POR_CAMPO[campo], valor)
            elif campo.endswith('_nd'):
                valor = 'Sí' if valor else ''
            if isinstance(valor, (list, tuple)):
                valor = '; '.join(str(x) for x in valor)
            fila.append(valor)
        v04.append(fila)
    for cell in v04[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='17365D')
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    v04.freeze_panes = 'C2'
    v04.auto_filter.ref = v04.dimensions
    v04.row_dimensions[1].height = 60
    for column_cells in v04.columns:
        max_length = min(max(len(str(cell.value or '')) for cell in column_cells) + 2, 42)
        v04.column_dimensions[get_column_letter(column_cells[0].column)].width = max(max_length, 12)
    wb.active = 0
    ws.title = 'Formulario anterior'

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    nombre_archivo = f"diagnostico_archivistico_matriz_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=nombre_archivo,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


# ---------------------------------------------------------------------------
# MÓDULO NUEVO: Sistema Distrital de Archivos (tablero público).
# Solo lee los diagnósticos existentes; no crea tablas ni modifica registros.
# ---------------------------------------------------------------------------
# Mínimo de entidades para publicar indicadores agregados sin exponer a una
# entidad individual. El panel de administración siempre ve los datos.
SDA_MIN_ENTIDADES_PUBLICO = int(os.environ.get('SDA_MIN_ENTIDADES_PUBLICO', '3'))

SDA_COMPONENTES = [
    ('estrategico', 'Estratégico'),
    ('administracion_archivos', 'Administración de archivos'),
    ('procesos_documentales', 'Procesos de gestión documental'),
    ('tecnologico', 'Tecnológico'),
    ('cultural', 'Cultural'),
]

# Estado de las TRD en el orden del ciclo que acompaña el Consejo (pregunta 25, hija «Estado de la TRD»).
SDA_ETAPAS_TRD = [
    ('No cuenta', 'Sin TRD'),
    ('Elaborada', 'Elaborada'),
    ('Aprobada', 'Aprobada'),
    ('Convalidada', 'Convalidada'),
    ('Implementada', 'Implementada'),
]

# Temas de normalización: (clave, tema, instrumento o lineamiento sugerido, pregunta del cuestionario).
# Hay brecha cuando la entidad responde «No» a la pregunta.
SDA_TEMAS_NORMALIZACION = [
    ('politica', 'Política de Gestión Documental', 'Guía para formular, aprobar y hacer seguimiento a la política', 'p2'),
    ('pgd', 'Programa de Gestión Documental', 'Lineamiento para elaborar y ejecutar el PGD', 'p5'),
    ('pinar', 'Plan Institucional de Archivos', 'Modelo de PINAR y su actualización anual', 'p6'),
    ('depositos', 'Depósitos y condiciones de almacenamiento', 'Lineamiento de condiciones de depósitos', 'p16'),
    ('trd', 'Tablas de Retención y Valoración Documental', 'Guía de elaboración y presentación de TRD y TVD', 'p25'),
    ('sic', 'Sistema Integrado de Conservación', 'Guía del Plan de Conservación y del Plan de Preservación Digital', 'p40'),
    ('sgdea', 'Documento electrónico y SGDEA', 'Lineamiento de requisitos para SGDEA', 'p44'),
    ('preservacion', 'Preservación digital', 'Guía del Plan de Preservación Digital a Largo Plazo', 'p50'),
    ('capacitacion', 'Capacitación en gestión documental', 'Plan distrital de capacitación en gestión documental', 'p21'),
    ('patrimonio', 'Patrimonio documental y archivos de DDHH', 'Lineamiento de identificación de fondos patrimoniales y de DDHH', 'p57'),
]

# Procesos de la gestión documental: subcomponentes del componente 3 del cuestionario.
SDA_PROCESOS = [sub for comp in cq.COMPONENTES if comp['clave'] == 'procesos_documentales'
                for sub in comp['subcomponentes']]


def _sda_lista(valor):
    if isinstance(valor, list):
        return valor
    return [valor] if valor else []


def _sda_pct(parte, total):
    return round(parte * 100 / total, 1) if total else 0


def _sda_nombre(registro):
    return (registro.nombre_entidad or '').split(' | NIT')[0].strip() or 'Entidad sin nombre'


def indicadores_sistema_distrital():
    """Agrega los diagnósticos más recientes de cada entidad en indicadores por función."""
    registros = Censo.query.order_by(Censo.fecha.desc()).all()
    vistos, entidades = set(), []
    for r in registros:  # se conserva solo el diagnóstico más reciente por entidad
        clave = (r.codigo_entidad or '').strip() or clave_entidad(_sda_nombre(r))
        if clave in vistos:
            continue
        vistos.add(clave)
        try:
            respuestas = json.loads(r.respuestas_json or '{}')
        except (TypeError, ValueError):
            respuestas = {}
        if not isinstance(respuestas, dict):
            respuestas = {}
        entidades.append({
            'id': r.id, 'nombre': _sda_nombre(r), 'sector': (r.sector or 'Sin sector').strip() or 'Sin sector',
            'ima': round(r.ima or 0, 2), 'fecha': r.fecha, 'respuestas': respuestas,
            'componentes': puntajes_componentes(respuestas),
        })

    n = len(entidades)
    try:
        en_curso = BorradorDiagnostico.query.filter_by(estado='borrador').count()
    except Exception:
        en_curso = 0

    def nivel(ima):
        return 'Bajo' if ima <= UMBRAL_BAJO else ('Medio' if ima <= UMBRAL_MEDIO else 'Alto')

    niveles = {'Bajo': 0, 'Medio': 0, 'Alto': 0}
    for e in entidades:
        niveles[nivel(e['ima'])] += 1

    # SEGUIMIENTO: promedio por componente y por sector.
    componentes = []
    for clave, etiqueta in SDA_COMPONENTES:
        valores = [e['componentes'].get(clave) for e in entidades if e['componentes'].get(clave) is not None]
        criticas = sum(1 for v in valores if v < 0.5)
        componentes.append({
            'clave': clave, 'etiqueta': etiqueta, 'respondidas': len(valores),
            'promedio': round(sum(valores) / len(valores) * 100, 1) if valores else None,
            'brecha_critica': criticas, 'brecha_pct': _sda_pct(criticas, len(valores)),
        })
    sectores = {}
    for e in entidades:
        s = sectores.setdefault(e['sector'], {'sector': e['sector'], 'entidades': 0, 'suma': 0.0})
        s['entidades'] += 1
        s['suma'] += e['ima']
    sectores = sorted(({'sector': s['sector'], 'entidades': s['entidades'],
                        'promedio': round(s['suma'] / s['entidades'], 1)} for s in sectores.values()),
                      key=lambda s: (-s['entidades'], s['sector']))

    # ASISTENCIA TÉCNICA: prioridad de acompañamiento y componente más débil.
    prioridades = {'Alta': [], 'Media': [], 'Baja': []}
    for e in entidades:
        p = {'Bajo': 'Alta', 'Medio': 'Media', 'Alto': 'Baja'}[nivel(e['ima'])]
        puntuados = [(v, et) for (c, et) in SDA_COMPONENTES
                     for v in [e['componentes'].get(c)] if v is not None]
        debil = min(puntuados)[1] if puntuados else 'Sin información suficiente'
        prioridades[p].append({'id': e['id'], 'nombre': e['nombre'], 'ima': e['ima'],
                               'sector': e['sector'], 'componente_debil': debil})
    for lista in prioridades.values():
        lista.sort(key=lambda x: x['ima'])
    focos = sorted([c for c in componentes if c['brecha_critica']], key=lambda c: -c['brecha_pct'])

    # NORMALIZACIÓN: temas con mayor proporción de entidades con brecha.
    # Solo cuentan los diagnósticos del cuestionario vigente (V04).
    con_dato = [e for e in entidades if cq.es_v04(e['respuestas'])]
    temas = []
    for clave, tema, instrumento, campo in SDA_TEMAS_NORMALIZACION:
        brecha = sum(1 for e in con_dato if e['respuestas'].get(campo) == 'no')
        temas.append({'clave': clave, 'tema': tema, 'instrumento': instrumento,
                      'entidades': brecha, 'pct': _sda_pct(brecha, len(con_dato))})
    temas.sort(key=lambda t: -t['pct'])
    procesos = []
    for sub in SDA_PROCESOS:
        valores = [v for e in con_dato for p in sub['preguntas']
                   for v in [cq.puntaje_pregunta(p, e['respuestas'])] if v is not None]
        procesos.append({'proceso': sub['nombre'], 'entidades': len(con_dato),
                         'pct': round(sum(valores) / len(valores) * 100, 1) if valores else 0})
    procesos.sort(key=lambda p: p['pct'])

    # EVALUACIÓN: ciclo de las TRD (pregunta 25).
    estados = {}
    for e in con_dato:
        estado = e['respuestas'].get('p25_trd') or 'Sin dato'
        estados[estado] = estados.get(estado, 0) + 1
    etapas = [{'estado': est, 'etiqueta': et, 'entidades': estados.get(est, 0),
               'pct': _sda_pct(estados.get(est, 0), len(con_dato))} for est, et in SDA_ETAPAS_TRD]
    otros_trd = {k: v for k, v in estados.items() if k not in dict(SDA_ETAPAS_TRD)}
    con_trd_avanzada = [e for e in con_dato if e['respuestas'].get('p25_trd') in ('Convalidada', 'Implementada')]
    con_evidencia = sum(1 for e in con_trd_avanzada if str(e['respuestas'].get('p25_url', '')).startswith('http'))
    requieren_actualizacion = sum(1 for e in con_dato if e['respuestas'].get('p25') in ('no', 'parcial'))

    return {
        'n': n, 'en_curso': en_curso, 'diagnosticos_totales': len(registros),
        'promedio': round(sum(e['ima'] for e in entidades) / n, 1) if n else None,
        'niveles': niveles, 'niveles_pct': {k: _sda_pct(v, n) for k, v in niveles.items()},
        'componentes': componentes, 'sectores': sectores,
        'prioridades': prioridades, 'focos': focos,
        'temas': temas, 'procesos': procesos, 'n_formulario_actual': len(con_dato),
        'etapas_trd': etapas, 'otros_trd': otros_trd,
        'trd_avanzada': len(con_trd_avanzada), 'trd_con_evidencia': con_evidencia,
        'trd_actualizacion': requieren_actualizacion,
        'actualizado': max((e['fecha'] for e in entidades), default=None),
    }


# ----- Indicadores del Sistema: solo administración -----
@app.route('/admin/sistema-distrital')
@admin_required
def admin_sistema_distrital():
    try:
        datos = indicadores_sistema_distrital()
    except Exception:
        app.logger.exception('No fue posible calcular los indicadores del Sistema Distrital')
        datos = None
    return render_template('admin_sistema_distrital.html', datos=datos, publicable=bool(datos),
                           es_admin=True, minimo=0)


@app.route('/api/admin/sistema-distrital/indicadores')
@admin_required
def api_sistema_distrital():
    """Indicadores agregados en JSON para uso administrativo (sin nombres de entidades)."""
    datos = indicadores_sistema_distrital()
    salida = {k: v for k, v in datos.items() if k not in ('prioridades', 'actualizado')}
    salida['prioridades'] = {k: len(v) for k, v in datos['prioridades'].items()}
    salida['actualizado'] = datos['actualizado'].isoformat() if datos['actualizado'] else None
    return jsonify(ok=True, **salida)


# ---------------------------------------------------------------------------
# Servicios en línea del Sistema Distrital: solicitudes de asistencia técnica y
# radicación de TRD/TVD para evaluación, con seguimiento por radicado.
# Tablas nuevas (se crean solas con db.create_all; no alteran las existentes).
# ---------------------------------------------------------------------------
class SolicitudSistema(db.Model):
    __tablename__ = 'solicitud_sistema'
    id = db.Column(db.Integer, primary_key=True)
    radicado = db.Column(db.String(40), unique=True, index=True)
    tipo = db.Column(db.String(20), nullable=False)          # asistencia | evaluacion
    entidad = db.Column(db.String(300), nullable=False)
    nit = db.Column(db.String(40), nullable=False)
    contacto_nombre = db.Column(db.String(200), nullable=False)
    contacto_cargo = db.Column(db.String(200))
    correo = db.Column(db.String(200), nullable=False)
    telefono = db.Column(db.String(60))
    tema = db.Column(db.String(200), nullable=False)          # tema de asistencia o instrumento
    modalidad = db.Column(db.String(60))
    enlace_documentos = db.Column(db.String(600))
    descripcion = db.Column(db.Text, nullable=False)
    estado = db.Column(db.String(60), nullable=False)
    nota_interna = db.Column(db.Text)
    creado_en = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    actualizado_en = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)


class SolicitudEvento(db.Model):
    __tablename__ = 'solicitud_evento'
    id = db.Column(db.Integer, primary_key=True)
    solicitud_id = db.Column(db.Integer, db.ForeignKey('solicitud_sistema.id'), nullable=False, index=True)
    fecha = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    estado = db.Column(db.String(60), nullable=False)
    observacion = db.Column(db.Text)                          # visible para la entidad


SDA_TIPOS = {
    'asistencia': {
        'prefijo': 'AT', 'nombre': 'Asistencia técnica',
        'temas': ['Política y Programa de Gestión Documental', 'Tablas de Retención y Valoración Documental',
                  'Organización de archivos', 'Conservación y depósitos', 'Documento electrónico y SGDEA',
                  'Preservación digital', 'Transferencias documentales', 'Capacitación', 'Otro'],
        'modalidades': ['Virtual', 'Presencial', 'Mesa técnica'],
        'estados': ['Recibida', 'Asignada', 'En atención', 'Atendida', 'Cerrada'],
    },
    'evaluacion': {
        'prefijo': 'EV', 'nombre': 'Evaluación de instrumentos',
        'temas': ['TRD – presentación por primera vez', 'TRD – actualización', 'TVD'],
        'modalidades': [],
        'estados': ['Recibida', 'Revisión preliminar', 'Mesa técnica', 'Requiere ajustes',
                    'Para sesión del Consejo', 'Concepto emitido', 'Cerrada'],
    },
}
SDA_MAX_SOLICITUDES_SESION = 5


@app.template_filter('hora_col')
def hora_colombia(valor, formato='%d/%m/%Y %H:%M'):
    """Fechas guardadas en UTC mostradas en hora de Colombia (UTC-5, sin horario de verano)."""
    if not valor:
        return ''
    return (valor - timedelta(hours=5)).strftime(formato)
# Anillo decorativo de entidades en el diagrama de la página pública (no son datos).
app.jinja_env.globals['sda_anillo'] = [
    (round(320 + 285 * math.cos(i * 2 * math.pi / 44), 1), round(320 + 285 * math.sin(i * 2 * math.pi / 44), 1),
     5 if i % 3 == 0 else 3.5, '#f1c75b' if i % 7 == 0 else '#8fc1f0', round((i * 0.37) % 4, 2))
    for i in range(44)]
_CORREO_RE = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')


def nit_base(nit):
    """Dígitos del NIT sin el dígito de verificación (890905211-1 -> 890905211)."""
    return re.sub(r'\D', '', (nit or '').split('-')[0])


def generar_radicado(prefijo, nit):
    """Radicado del Consejo Distrital de Archivos: CDA-<tipo>-<NIT>-<año>-<consecutivo>.
    El consecutivo es independiente por tipo de solicitud y se reinicia cada año."""
    anio = datetime.utcnow().year
    patron = f'CDA-{prefijo}-%-{anio}-%'
    n = SolicitudSistema.query.filter(SolicitudSistema.radicado.like(patron)).count() + 1
    while True:
        radicado = f'CDA-{prefijo}-{nit_base(nit)}-{anio}-{n:04d}'
        if not SolicitudSistema.query.filter_by(radicado=radicado).first():
            return radicado
        n += 1


def _sda_radicados_sesion():
    return list(session.get('sda_radicados', []))


@app.route('/sistema-distrital')
def sistema_distrital():
    # La información del Sistema y del Consejo ahora es la página de inicio.
    return redirect(url_for('index'), code=301)


# Servicios que las entidades pueden solicitar en línea. La evaluación de TRD y TVD
# no se tramita por la herramienta: se presenta por el canal de correspondencia.
SERVICIOS_PUBLICOS = ('asistencia',)


@app.route('/servicios-en-linea')
def servicios_en_linea():
    servicio = request.args.get('servicio', '')
    return render_template('servicios_en_linea.html', tipos=SDA_TIPOS, valores={}, errores={},
                           tipo_activo=servicio if servicio in SERVICIOS_PUBLICOS + ('consulta',) else '')


@app.route('/sistema-distrital/solicitud', methods=['POST'])
def sistema_distrital_solicitud():
    f = request.form
    tipo = f.get('tipo', '')
    if tipo not in SERVICIOS_PUBLICOS:
        abort(400)
    if f.get('sitio_web_alterno'):          # campo trampa contra robots
        return redirect(url_for('sistema_distrital'))
    cfg = SDA_TIPOS[tipo]
    v = {k: (f.get(k, '') or '').strip() for k in
         ['entidad', 'nit', 'contacto_nombre', 'contacto_cargo', 'correo', 'telefono',
          'tema', 'modalidad', 'enlace_documentos', 'descripcion']}
    errores = {}
    for campo, etiqueta in [('entidad', 'el nombre de la entidad'), ('nit', 'el NIT'),
                            ('contacto_nombre', 'el nombre de contacto'), ('correo', 'el correo'),
                            ('descripcion', 'la descripción')]:
        if not v[campo]:
            errores[campo] = f'Ingrese {etiqueta}.'
    if v['correo'] and not _CORREO_RE.match(v['correo']):
        errores['correo'] = 'Ingrese un correo válido.'
    if v['tema'] not in cfg['temas']:
        errores['tema'] = 'Seleccione una opción.'
    if cfg['modalidades'] and v['modalidad'] not in cfg['modalidades']:
        errores['modalidad'] = 'Seleccione una modalidad.'
    if tipo == 'evaluacion':
        if not re.match(r'^https://\S+$', v['enlace_documentos']):
            errores['enlace_documentos'] = 'Ingrese un enlace https a la carpeta con los documentos.'
    elif v['enlace_documentos'] and not re.match(r'^https?://\S+$', v['enlace_documentos']):
        errores['enlace_documentos'] = 'El enlace debe comenzar por https://'
    if v['nit'] and not 6 <= len(nit_base(v['nit'])) <= 10:
        errores['nit'] = 'Ingrese un NIT válido, por ejemplo 890905211-1.'
    if not f.get('autoriza_datos'):
        errores['autoriza_datos'] = 'Debe autorizar el tratamiento de datos para enviar la solicitud.'
    limites = {'entidad': 300, 'nit': 40, 'contacto_nombre': 200, 'contacto_cargo': 200, 'correo': 200,
               'telefono': 60, 'enlace_documentos': 600, 'descripcion': 4000}
    for campo, maximo in limites.items():
        if len(v[campo]) > maximo:
            errores[campo] = f'Máximo {maximo} caracteres.'
    if len(_sda_radicados_sesion()) >= SDA_MAX_SOLICITUDES_SESION:
        errores['general'] = 'Se alcanzó el número máximo de solicitudes desde este navegador. Intente más tarde.'
    if errores:
        return render_template('servicios_en_linea.html', tipos=SDA_TIPOS,
                               valores=v, errores=errores, tipo_activo=tipo), 400
    try:
        s = SolicitudSistema(tipo=tipo, entidad=v['entidad'], nit=v['nit'], contacto_nombre=v['contacto_nombre'],
                             contacto_cargo=v['contacto_cargo'], correo=v['correo'].lower(), telefono=v['telefono'],
                             tema=v['tema'], modalidad=v['modalidad'] or None,
                             enlace_documentos=v['enlace_documentos'] or None,
                             descripcion=v['descripcion'], estado='Recibida')
        db.session.add(s)
        db.session.flush()
        s.radicado = generar_radicado(cfg['prefijo'], v['nit'])
        db.session.add(SolicitudEvento(solicitud_id=s.id, estado='Recibida',
                                       observacion='Solicitud recibida por el Consejo Distrital de Archivos.'))
        db.session.commit()
    except Exception:
        db.session.rollback()
        app.logger.exception('No fue posible radicar la solicitud del Sistema Distrital')
        errores['general'] = 'No fue posible registrar la solicitud en este momento. Intente de nuevo.'
        return render_template('servicios_en_linea.html', tipos=SDA_TIPOS,
                               valores=v, errores=errores, tipo_activo=tipo), 500
    session['sda_radicados'] = (_sda_radicados_sesion() + [s.radicado])[-20:]
    return redirect(url_for('sistema_distrital_tramite', radicado=s.radicado, nuevo=1))


@app.route('/sistema-distrital/tramite', methods=['GET', 'POST'])
def sistema_distrital_tramite():
    radicado = (request.values.get('radicado', '') or '').strip().upper()
    correo = (request.form.get('correo', '') or '').strip().lower()
    solicitud, error = None, None
    if radicado and (request.method == 'POST' or radicado in _sda_radicados_sesion()):
        candidata = SolicitudSistema.query.filter_by(radicado=radicado).first()
        # Se exige el correo registrado, salvo para quien acaba de radicar desde este navegador.
        if candidata and (radicado in _sda_radicados_sesion() or candidata.correo == correo):
            solicitud = candidata
        else:
            error = 'No encontramos una solicitud con ese número y correo. Verifique los datos.'
    eventos = []
    if solicitud:
        eventos = SolicitudEvento.query.filter_by(solicitud_id=solicitud.id).order_by(SolicitudEvento.fecha.asc()).all()
    return render_template('sistema_distrital_tramite.html', solicitud=solicitud, eventos=eventos,
                           tipos=SDA_TIPOS, radicado=radicado, error=error,
                           nuevo=bool(request.args.get('nuevo')) and solicitud is not None)


@app.route('/admin/solicitudes')
@admin_required
def admin_solicitudes():
    tipo = request.args.get('tipo', '')
    estado = request.args.get('estado', '')
    q = SolicitudSistema.query
    if tipo in SDA_TIPOS:
        q = q.filter_by(tipo=tipo)
    if estado:
        q = q.filter_by(estado=estado)
    solicitudes = q.order_by(SolicitudSistema.creado_en.desc()).all()
    eventos = {}
    if solicitudes:
        for e in SolicitudEvento.query.filter(SolicitudEvento.solicitud_id.in_([s.id for s in solicitudes])) \
                .order_by(SolicitudEvento.fecha.asc()).all():
            eventos.setdefault(e.solicitud_id, []).append(e)
    abiertas = SolicitudSistema.query.filter(SolicitudSistema.estado.notin_(['Cerrada'])).count()
    return render_template('admin_solicitudes.html', solicitudes=solicitudes, eventos=eventos, tipos=SDA_TIPOS,
                           filtro_tipo=tipo, filtro_estado=estado, abiertas=abiertas)


@app.route('/admin/solicitudes/<int:solicitud_id>', methods=['POST'])
@admin_required
def admin_actualizar_solicitud(solicitud_id):
    s = db.session.get(SolicitudSistema, solicitud_id) or abort(404)
    nuevo_estado = request.form.get('estado', '')
    observacion = (request.form.get('observacion', '') or '').strip()[:4000]
    s.nota_interna = (request.form.get('nota_interna', '') or '').strip()[:4000] or None
    if nuevo_estado not in SDA_TIPOS[s.tipo]['estados']:
        flash('Estado no válido.', 'danger')
        return redirect(url_for('admin_solicitudes'))
    try:
        if nuevo_estado != s.estado or observacion:
            db.session.add(SolicitudEvento(solicitud_id=s.id, estado=nuevo_estado, observacion=observacion or None))
            s.estado = nuevo_estado
        db.session.commit()
        flash(f'Solicitud {s.radicado} actualizada.', 'success')
    except Exception:
        db.session.rollback()
        app.logger.exception('No fue posible actualizar la solicitud')
        flash('No fue posible actualizar la solicitud.', 'danger')
    return redirect(url_for('admin_solicitudes', tipo=request.args.get('tipo', ''),
                            estado=request.args.get('estado', '')) + f'#sol-{s.id}')

# ---------------------------------------------------------------------------
# Ficha de diagnóstico por entidad (administración y la propia entidad).
# Muestra el resultado por componente y cada respuesta con sus preguntas hijas y
# evidencias. Solo lee datos; no modifica registros.
# ---------------------------------------------------------------------------
FICHA_PROCESOS = ['Planeación', 'Producción', 'Gestión y trámite', 'Organización', 'Transferencias',
                  'Disposición final', 'Preservación a largo plazo', 'Valoración']
FICHA_PROTECCIONES = ['Copias de seguridad verificadas', 'Controles de acceso', 'Registros de auditoría',
                      'Verificación de integridad', 'Migración de formatos', 'Plan de preservación digital']

# Componentes de la ficha para los diagnósticos del formulario anterior (secciones A–F).
FICHA_COMPONENTES_ANTERIOR = [
    {'clave': 'identificacion', 'letra': 'A', 'nombre': 'Identificación y caracterización', 'componente': None,
     'preguntas': [
         ('Responsable de gestión documental', ['responsable_nombre', 'responsable_cargo', 'responsable_correo', 'telefono_responsable']),
         ('Naturaleza jurídica', ['naturaleza']), ('Sector', ['sector']),
         ('Grupo / relación con el Distrito', ['grupo_relacion']), ('Clasificación preliminar', ['clasificacion_preliminar']),
         ('Ámbito', ['ambito']), ('Dirección', ['direccion_entidad']), ('Sitio web', ['sitio_web']),
         ('Datos corregidos frente al directorio', ['correcciones_directorio']),
     ]},
    {'clave': 'estrategico', 'letra': 'B', 'nombre': 'Estratégico', 'componente': 'estrategico',
     'preguntas': [
         ('1. Estado de la Política de Gestión Documental', ['q1_politica', 'q1_acto', 'q1_acto_detalle', 'q1_etapa', 'q1_razon', 'evidencia_q1']),
         ('2. Seguimiento al PGD y acciones de mejora', ['q2_seguimiento', 'q2_mecanismo', 'q2_pgd_formulado', 'evidencia_q2']),
     ]},
    {'clave': 'administracion_archivos', 'letra': 'C', 'nombre': 'Administración de archivos', 'componente': 'administracion_archivos',
     'preguntas': [
         ('3. Personas en gestión documental y archivo', ['q3_personal', 'q3_observaciones']),
         ('4. Condiciones de los depósitos de archivo', ['q4_depositos', 'q4_riesgos', 'evidencia_q4']),
     ]},
    {'clave': 'procesos_documentales', 'letra': 'D', 'nombre': 'Procesos de gestión documental', 'componente': 'procesos_documentales',
     'preguntas': [
         ('5. Estado de las Tablas de Retención Documental', ['q5_trd', 'q5_avance', 'q5_evaluacion', 'q5_evidencia_estado', 'q5_actualizacion', 'evidencia_q5']),
         ('6. Procesos con procedimientos documentados', ['q6_procesos', 'evidencia_q6']),
     ]},
    {'clave': 'tecnologico', 'letra': 'E', 'nombre': 'Tecnológico', 'componente': 'tecnologico',
     'preguntas': [
         ('7. SGDEA o sistema equivalente', ['q7_sgdea', 'q7_sistema_detalle', 'q7_soporte', 'q7_limitacion', 'evidencia_q7']),
         ('8. Seguridad y preservación digital', ['q8_protecciones', 'evidencia_q8']),
     ]},
    {'clave': 'cultural', 'letra': 'F', 'nombre': 'Cultural', 'componente': 'cultural',
     'preguntas': [
         ('9. Capacitación y sensibilización', ['q9_capacitacion', 'q9_detalle', 'q9_barreras', 'evidencia_q9']),
         ('10. Acción prioritaria para el acceso y uso de la información', ['q10_prioridad', 'evidencia_q10']),
     ]},
]

FICHA_ETIQUETAS = {
    'responsable_nombre': 'Nombre', 'responsable_cargo': 'Cargo', 'responsable_correo': 'Correo',
    'telefono_responsable': 'Teléfono', 'correcciones_directorio': 'Datos corregidos frente al directorio', 'q1_acto': 'Acto administrativo', 'q1_acto_detalle': 'Número y fecha del acto',
    'q1_etapa': 'Etapa', 'q1_razon': 'Razón por la que no existe', 'q2_mecanismo': 'Mecanismo de seguimiento',
    'q2_pgd_formulado': 'PGD formulado', 'q3_observaciones': 'Observaciones', 'q4_riesgos': 'Deficiencias identificadas',
    'q5_avance': 'Avance pendiente', 'q5_evaluacion': 'Instancia y fecha de evaluación',
    'q5_evidencia_estado': 'Constancia de convalidación e implementación', 'q5_actualizacion': 'Motivo de actualización',
    'q7_sistema_detalle': 'Sistema y procesos que cubre', 'q7_soporte': 'Documentación o soporte',
    'q7_limitacion': 'Principal limitación', 'q9_detalle': 'Temas, población y frecuencia', 'q9_barreras': 'Barreras',
}

def _ficha_valor(r, campo):
    valor = r.get(campo, '')
    if isinstance(valor, list):
        return [str(v).strip() for v in valor if str(v).strip()]
    return str(valor or '').strip()


def _ficha_clave(registro):
    return (registro.codigo_entidad or '').strip() or clave_entidad(_sda_nombre(registro))


def _puntajes_de(censo):
    """Respuestas y puntaje por componente de un diagnóstico guardado."""
    try:
        respuestas = json.loads(censo.respuestas_json or '{}')
    except (TypeError, ValueError):
        respuestas = {}
    if not isinstance(respuestas, dict):
        respuestas = {}
    actual = cq.es_v04(respuestas) or bool(respuestas.get('q1_politica'))
    return respuestas, actual, (puntajes_componentes(respuestas) if actual else {})


def _version_de(respuestas):
    return cq.VERSION if cq.es_v04(respuestas) else 'anterior'


def _nivel_de(pct):
    if pct is None:
        return 'Sin respuestas puntuables'
    return 'Bajo' if pct <= UMBRAL_BAJO else ('Medio' if pct <= UMBRAL_MEDIO else 'Alto')


def tabla_componentes(respuestas, componentes):
    """Filas de la tabla de resultados: una por componente con su porcentaje."""
    filas = []
    if cq.es_v04(respuestas):
        for comp in cq.COMPONENTES:
            preguntas = [p for p in cq.PREGUNTAS if p['componente'] == comp['clave']]
            puntos = [v for p in preguntas for v in [cq.puntaje_pregunta(p, respuestas)] if v is not None]
            v = componentes.get(comp['clave'])
            pct = None if v is None else round(v * 100, 1)
            filas.append({'num': comp['num'], 'nombre': comp['nombre'], 'peso': comp['peso'],
                          'preguntas': len(preguntas), 'aplicables': len(puntos),
                          'puntos': round(sum(puntos), 1), 'pct': pct, 'nivel': _nivel_de(pct),
                          'si': sum(1 for p in preguntas if respuestas.get(p['campo']) == 'si'),
                          'parcial': sum(1 for p in preguntas if respuestas.get(p['campo']) == 'parcial'),
                          'no': sum(1 for p in preguntas if respuestas.get(p['campo']) == 'no'),
                          'na': len(preguntas) - len(puntos) - 0})
        for f in filas:  # N/A y excluidas = preguntas que no cuentan en el puntaje
            f['na'] = f['preguntas'] - f['aplicables']
    else:
        for n, (clave, nombre) in enumerate(SDA_COMPONENTES, start=1):
            v = componentes.get(clave)
            pct = None if v is None else round(v * 100, 1)
            filas.append({'num': n, 'nombre': nombre, 'peso': 20, 'preguntas': None, 'aplicables': None,
                          'puntos': None, 'pct': pct, 'nivel': _nivel_de(pct),
                          'si': None, 'parcial': None, 'no': None, 'na': None})
    return filas


def _valor_hija(h, respuestas):
    """Valor legible de una pregunta hija para la ficha."""
    if h['tipo'] == 'matriz':
        base = h['campo']
        filas = [f"{fila}: {respuestas.get(base + '_' + str(i))}" for i, fila in enumerate(h['filas'])
                 if respuestas.get(base + '_' + str(i))]
        return '\n'.join(filas)
    valor = _ficha_valor(respuestas, h['campo'])
    if h['tipo'] == 'multi':
        valor = valor if isinstance(valor, list) else ([valor] if valor else [])
        otro = _ficha_valor(respuestas, f"{h['campo']}_otro")
        if otro:
            valor = [f"{v}: {otro}" if v == h.get('otro') else v for v in valor]
    if h['tipo'] == 'url' and not valor and respuestas.get(f"{h['campo']}_nd"):
        valor = 'Enlace no disponible (para revisión)'
    return valor


def _componentes_ficha_v04(respuestas, componentes, componentes_previos):
    r = respuestas
    fechas = 'Fecha desconocida' if r.get('fechas_desconocidas') else ' – '.join(
        x for x in [_ficha_valor(r, 'fecha_inicial'), _ficha_valor(r, 'fecha_final')] if x)
    anio = re.search(r'(\d{4})', _ficha_valor(r, 'fecha_inicial'))
    caracterizacion = [
        ('C1. Naturaleza jurídica', _ficha_valor(r, 'naturaleza'), []),
        ('C2. Nombre de la entidad', _ficha_valor(r, 'nombre_entidad'), [('NIT', _ficha_valor(r, 'nit_entidad'))]),
        ('C3. Sector administrativo', _ficha_valor(r, 'sector'), [('¿Cuál?', _ficha_valor(r, 'sector_otro'))]),
        ('C4. Dirección de la sede principal', _ficha_valor(r, 'direccion_entidad'), []),
        ('C5. Teléfono institucional', _ficha_valor(r, 'telefono_entidad'), []),
        ('C6. Sitio web institucional', _ficha_valor(r, 'sitio_web'), []),
        ('C7. Representante legal', _ficha_valor(r, 'rep_nombre'),
         [('Cargo', _ficha_valor(r, 'rep_cargo')), ('Correo', _ficha_valor(r, 'rep_correo')), ('Teléfono', _ficha_valor(r, 'rep_telefono'))]),
        ('C8. Responsable de gestión documental y archivos', _ficha_valor(r, 'responsable_nombre'),
         [('Cargo', _ficha_valor(r, 'responsable_cargo')), ('Dependencia', _ficha_valor(r, 'responsable_dependencia')),
          ('Correo', _ficha_valor(r, 'responsable_correo')), ('Teléfono', _ficha_valor(r, 'telefono_responsable'))]),
        ('C9. Fechas extremas de los documentos', fechas,
         [('Observaciones', _ficha_valor(r, 'fechas_observaciones')),
          ('Alerta', 'Documentos anteriores a 1911: revisar posible interés patrimonial (BIC-CDA).'
           if anio and int(anio.group(1)) < 1911 else '')]),
        ('Datos corregidos frente al directorio', _ficha_valor(r, 'correcciones_directorio'), []),
    ]
    secciones = [{'clave': 'identificacion', 'letra': 'C', 'nombre': 'Identificación y caracterización institucional',
             'componente': None, 'puntaje': None, 'anterior': None, 'cambio': None, 'extra': None,
             'preguntas': [{'titulo': t, 'sub': None, 'valor': v, 'etiqueta_principal': None,
                            'detalles': [{'etiqueta': e, 'valor': x, 'es_evidencia': False} for e, x in det if x]}
                           for t, v, det in caracterizacion]}]
    for comp in cq.COMPONENTES:
        preguntas = []
        for sub in comp['subcomponentes']:
            for n, p in enumerate(sub['preguntas']):
                detalles = []
                for h in p['hijas']:
                    if not cq.hija_visible(h, r, p):
                        continue
                    v = _valor_hija(h, r)
                    if v:
                        detalles.append({'etiqueta': 'Evidencia' if h['tipo'] == 'url' else h['etiqueta'],
                                         'valor': v, 'es_evidencia': h['tipo'] == 'url'})
                puntos = cq.puntaje_pregunta(p, r)
                preguntas.append({'titulo': f"{p['n']}. {p['texto']}", 'sub': sub['nombre'] if n == 0 else None,
                                  'valor': cq.etiqueta_opcion(p, r.get(p['campo'], '')),
                                  'puntos': puntos, 'etiqueta_principal': None, 'detalles': detalles})
        v, prev = componentes.get(comp['clave']), componentes_previos.get(comp['clave'])
        pct = None if v is None else round(v * 100, 1)
        pct_prev = None if prev is None else round(prev * 100, 1)
        secciones.append({'clave': comp['clave'], 'letra': str(comp['num']), 'nombre': comp['nombre'],
                     'componente': comp['clave'], 'puntaje': pct, 'anterior': pct_prev,
                     'cambio': None if pct is None or pct_prev is None else round(pct - pct_prev, 1),
                     'preguntas': preguntas, 'extra': None})
    return secciones


def _contexto_ficha(censo):
    """Datos de la ficha que comparten la vista del Consejo y la de la entidad."""
    respuestas, formulario_actual, componentes = _puntajes_de(censo)
    version = _version_de(respuestas)

    todos = Censo.query.order_by(Censo.fecha.desc(), Censo.id.desc()).all()
    clave = _ficha_clave(censo)
    historial = [x for x in todos if _ficha_clave(x) == clave]
    # Diagnóstico anterior con la misma versión del formulario, para comparar por componente.
    anterior, componentes_previos = None, {}
    posicion = next((n for n, x in enumerate(historial) if x.id == censo.id), None)
    if posicion is not None:
        for x in historial[posicion + 1:]:
            resp_x, actual_x, comp_x = _puntajes_de(x)
            if actual_x and _version_de(resp_x) == version:
                anterior, componentes_previos = x, comp_x
                break

    if version == cq.VERSION:
        componentes_ficha = _componentes_ficha_v04(respuestas, componentes, componentes_previos)
    else:
        componentes_ficha = _componentes_ficha_anteriores(respuestas, componentes, componentes_previos)
    return dict(censo=censo, nombre=_sda_nombre(censo), respuestas=respuestas, version=version,
                formulario_actual=formulario_actual, componentes_ficha=componentes_ficha,
                tabla_componentes=tabla_componentes(respuestas, componentes) if formulario_actual else [],
                historial=historial, nombre_de=_sda_nombre, anterior=anterior)


def _componentes_ficha_anteriores(respuestas, componentes, componentes_previos):
    """Ficha de los diagnósticos hechos con el formulario anterior (preguntas 1 a 10)."""
    componentes_ficha = []
    for d in FICHA_COMPONENTES_ANTERIOR:
        puntaje = componentes.get(d['componente']) if d['componente'] else None
        previo = componentes_previos.get(d['componente']) if d['componente'] else None
        preguntas = []
        for titulo, campos in d['preguntas']:
            principal = _ficha_valor(respuestas, campos[0])
            detalles = []
            for c in campos[1:]:
                v = _ficha_valor(respuestas, c)
                if v:
                    detalles.append({'campo': c, 'etiqueta': 'Evidencia' if c.startswith('evidencia_') else FICHA_ETIQUETAS.get(c, c),
                                     'valor': v, 'es_evidencia': c.startswith('evidencia_')})
            preguntas.append({'titulo': titulo, 'campo': campos[0], 'sub': None,
                              'etiqueta_principal': FICHA_ETIQUETAS.get(campos[0]) if d['clave'] == 'identificacion' and len(campos) > 1 else None,
                              'valor': principal, 'detalles': detalles})
        extra = None
        if d['clave'] == 'procesos_documentales':
            sel = set(_sda_lista(respuestas.get('q6_procesos')))
            extra = {'tipo': 'checklist', 'titulo': 'Procesos', 'elementos': [(p, p in sel) for p in FICHA_PROCESOS]}
        elif d['clave'] == 'tecnologico':
            sel = set(_sda_lista(respuestas.get('q8_protecciones')))
            extra = {'tipo': 'checklist', 'titulo': 'Mecanismos', 'elementos': [(p, p in sel) for p in FICHA_PROTECCIONES]}
        pct = None if puntaje is None else round(puntaje * 100, 1)
        pct_prev = None if previo is None else round(previo * 100, 1)
        componentes_ficha.append({**d, 'puntaje': pct, 'anterior': pct_prev,
                            'cambio': None if pct is None or pct_prev is None else round(pct - pct_prev, 1),
                            'preguntas': preguntas, 'extra': extra})
    return componentes_ficha


def _puede_ver_censo(censo):
    """El Consejo, quien envió el diagnóstico en esta sesión o la entidad dueña."""
    if session.get('admin_authenticated') or censo.id in session.get('mis_diagnosticos', []):
        return True
    entidad = entidad_en_sesion()
    return bool(entidad and censo in _diagnosticos_de(entidad))


@app.route('/admin/diagnostico/<int:censo_id>')
@admin_required
def ficha_diagnostico(censo_id):
    censo = db.session.get(Censo, censo_id) or abort(404)
    contexto = _contexto_ficha(censo)
    opciones = sorted(Censo.query.all(), key=lambda x: (_sda_nombre(x).lower(), -x.id))
    _, _, componentes = _puntajes_de(censo)
    acomp = {'lista': _acompanamientos_por_entidad().get(_ficha_clave(censo), []),
             'debil': _componente_mas_debil(componentes), 'tipo': _tipo_acompanamiento(componentes),
             'componentes': SDA_COMPONENTES, 'nombres': NOMBRE_COMPONENTE, 'hoy': date.today().isoformat()}
    return render_template('admin_ficha.html', vista='admin', opciones=opciones, acomp=acomp, **contexto)


@app.route('/mi-diagnostico/<int:censo_id>')
def ficha_entidad(censo_id):
    censo = db.session.get(Censo, censo_id) or abort(404)
    if not _puede_ver_censo(censo):
        if acceso_requerido() and not entidad_en_sesion():
            return redirect(url_for('acceso_entidad', next=request.path))
        abort(404)
    return render_template('admin_ficha.html', vista='entidad', opciones=[], **_contexto_ficha(censo))


# ============================================================
# TRANSPARENCIA, ACTAS Y CONVALIDACIONES
# Publicación de las actas y documentos del Consejo Distrital de Archivos y registro
# público de las TRD y TVD convalidadas (Ley 1712 de 2014, Res. MinTIC 1519 de 2020,
# Acuerdo AGN 001 de 2024). Los archivos PDF se guardan en la base de datos para que
# no se pierdan al desplegar en Render.
# ============================================================

PUB_CATEGORIAS = {
    'acta': {'nombre': 'Actas de sesión', 'singular': 'Acta', 'icono': 'bi-journal-check'},
    'plan_accion': {'nombre': 'Plan de acción anual', 'singular': 'Plan de acción', 'icono': 'bi-calendar2-check'},
    'informe_semestral': {'nombre': 'Informes semestrales al AGN', 'singular': 'Informe semestral', 'icono': 'bi-clipboard-data'},
    'reglamento': {'nombre': 'Reglamento interno y actos de creación', 'singular': 'Reglamento', 'icono': 'bi-bank'},
    'otro': {'nombre': 'Otros documentos', 'singular': 'Documento', 'icono': 'bi-file-earmark-text'},
}
PUB_TIPOS_SESION = ['Ordinaria', 'Extraordinaria']
PUB_MODALIDADES = ['Presencial', 'Virtual', 'Mixta']
CONV_INSTRUMENTOS = {'TRD': 'Tabla de Retención Documental', 'TVD': 'Tabla de Valoración Documental'}
CONV_TRAMITES = ['Primera vez', 'Actualización']
CONV_ESTADOS_RUSD = ['Pendiente de solicitud', 'Solicitada al AGN', 'Inscrita en el RUSD', 'Devuelta por el AGN']
PDF_MAX_BYTES = 10 * 1024 * 1024


class PublicacionCDA(db.Model):
    __tablename__ = 'publicacion_cda'
    id = db.Column(db.Integer, primary_key=True)
    categoria = db.Column(db.String(30), nullable=False, index=True)
    titulo = db.Column(db.String(250), nullable=False)
    numero = db.Column(db.Integer)
    anio = db.Column(db.Integer, index=True)
    fecha = db.Column(db.Date)
    tipo_sesion = db.Column(db.String(20))
    modalidad = db.Column(db.String(20))
    resumen = db.Column(db.Text)
    palabras_clave = db.Column(db.String(250))
    enlace = db.Column(db.String(500))
    archivo_nombre = db.Column(db.String(200))
    archivo_tamano = db.Column(db.Integer)
    archivo_hash = db.Column(db.String(64))
    archivo = deferred(db.Column(db.LargeBinary))
    publicado = db.Column(db.Boolean, default=True, nullable=False)
    fecha_publicacion = db.Column(db.DateTime)
    creado = db.Column(db.DateTime, default=datetime.utcnow)
    actualizado = db.Column(db.DateTime, default=datetime.utcnow)

    @property
    def nombre_acta(self):
        if self.categoria != 'acta':
            return self.titulo
        return f'Acta N.º {self.numero or 0:02d} de {self.anio}'


class ConvalidacionCDA(db.Model):
    __tablename__ = 'convalidacion_cda'
    id = db.Column(db.Integer, primary_key=True)
    entidad = db.Column(db.String(250), nullable=False)
    nit = db.Column(db.String(30), nullable=False)
    instrumento = db.Column(db.String(5), nullable=False)
    tramite = db.Column(db.String(30))
    version = db.Column(db.String(40))
    fecha_aprobacion_comite = db.Column(db.Date)
    fecha_convalidacion = db.Column(db.Date, nullable=False)
    acta_id = db.Column(db.Integer, db.ForeignKey('publicacion_cda.id'))
    acta = db.relationship('PublicacionCDA')
    numero_concepto = db.Column(db.String(60))
    radicado = db.Column(db.String(40))
    estado_rusd = db.Column(db.String(40), default='Pendiente de solicitud')
    fecha_rusd = db.Column(db.Date)
    observaciones = db.Column(db.Text)
    enlace = db.Column(db.String(500))
    archivo_nombre = db.Column(db.String(200))
    archivo_tamano = db.Column(db.Integer)
    archivo_hash = db.Column(db.String(64))
    archivo = deferred(db.Column(db.LargeBinary))
    publicado = db.Column(db.Boolean, default=True, nullable=False)
    creado = db.Column(db.DateTime, default=datetime.utcnow)
    actualizado = db.Column(db.DateTime, default=datetime.utcnow)


class AcompanamientoEntidad(db.Model):
    """Acompañamiento técnico del Consejo a una entidad priorizada (paso 5 del ciclo de la matriz)."""
    __tablename__ = 'acompanamiento_entidad'
    id = db.Column(db.Integer, primary_key=True)
    entidad_clave = db.Column(db.String(120), nullable=False, index=True)
    entidad_nombre = db.Column(db.String(300), nullable=False)
    censo_id = db.Column(db.Integer, nullable=False)              # diagnóstico con el que se inició
    nivel_inicial = db.Column(db.Float)                           # nivel de madurez de ese diagnóstico
    componente = db.Column(db.String(60), nullable=False)
    tipo = db.Column(db.String(20), nullable=False)               # Integral | Puntual
    responsable = db.Column(db.String(200), nullable=False)
    fecha_inicio = db.Column(db.Date, nullable=False)
    fecha_cierre = db.Column(db.Date)
    estado = db.Column(db.String(20), nullable=False, default='En curso')   # En curso | Cerrado
    observaciones = db.Column(db.Text)
    observacion_cierre = db.Column(db.Text)
    creado = db.Column(db.DateTime, default=datetime.utcnow)


def _fecha_form(valor):
    try:
        return datetime.strptime((valor or '').strip(), '%Y-%m-%d').date()
    except ValueError:
        return None


def _entero_form(valor):
    try:
        return int((valor or '').strip())
    except ValueError:
        return None


def _enlace_form(valor):
    valor = (valor or '').strip()[:500]
    return valor if valor.startswith(('https://', 'http://')) else ''


def _leer_pdf(campo='archivo'):
    """Devuelve (bytes, nombre, error). Solo acepta PDF reales de hasta 10 MB."""
    archivo = request.files.get(campo)
    if not archivo or not archivo.filename:
        return None, None, None
    datos = archivo.read(PDF_MAX_BYTES + 1)
    if len(datos) > PDF_MAX_BYTES:
        return None, None, 'El archivo supera 10 MB.'
    if not datos.startswith(b'%PDF-'):
        return None, None, 'El archivo debe ser un PDF.'
    nombre = re.sub(r'[^A-Za-z0-9._-]+', '_', unicodedata.normalize('NFKD', archivo.filename)
                    .encode('ascii', 'ignore').decode())[-150:] or 'documento.pdf'
    if not nombre.lower().endswith('.pdf'):
        nombre += '.pdf'
    return datos, nombre, None


def _asignar_pdf(obj, datos, nombre):
    obj.archivo, obj.archivo_nombre = datos, nombre
    obj.archivo_tamano, obj.archivo_hash = len(datos), hashlib.sha256(datos).hexdigest()


def _enviar_pdf(obj):
    if not obj.archivo_nombre:
        abort(404)
    respuesta = send_file(BytesIO(obj.archivo), mimetype='application/pdf',
                          download_name=obj.archivo_nombre, as_attachment=False)
    respuesta.headers['Cache-Control'] = 'public, max-age=3600'
    return respuesta


def _anios(columna):
    return sorted({a for (a,) in db.session.query(columna).distinct() if a}, reverse=True)


@app.template_filter('tamano')
def tamano_legible(n):
    if not n:
        return ''
    return f'{n / 1024:.0f} KB' if n < 1024 * 1024 else f'{n / 1024 / 1024:.1f} MB'


@app.template_filter('fecha_larga')
def fecha_larga(valor):
    if not valor:
        return ''
    meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto',
             'septiembre', 'octubre', 'noviembre', 'diciembre']
    return f'{valor.day} de {meses[valor.month - 1]} de {valor.year}'


def _actas_publicadas():
    return (PublicacionCDA.query.filter_by(categoria='acta', publicado=True)
            .order_by(PublicacionCDA.anio.desc(), PublicacionCDA.numero.desc()))


def _convalidaciones_publicadas():
    return (ConvalidacionCDA.query.filter_by(publicado=True)
            .order_by(ConvalidacionCDA.fecha_convalidacion.desc(), ConvalidacionCDA.id.desc()))


# ---------- Páginas públicas ----------

@app.route('/transparencia')
def transparencia():
    ultima = [x for x in (
        db.session.query(db.func.max(PublicacionCDA.fecha_publicacion)).filter(PublicacionCDA.publicado.is_(True)).scalar(),
        db.session.query(db.func.max(ConvalidacionCDA.actualizado)).filter(ConvalidacionCDA.publicado.is_(True)).scalar(),
    ) if x]
    conteos = {
        'actas': _actas_publicadas().count(),
        'convalidaciones': _convalidaciones_publicadas().count(),
        'documentos': PublicacionCDA.query.filter(PublicacionCDA.categoria != 'acta', PublicacionCDA.publicado.is_(True)).count(),
    }
    return render_template('transparencia.html', conteos=conteos, ultima=max(ultima) if ultima else None)


@app.route('/transparencia/actas')
def actas_publicas():
    consulta = _actas_publicadas()
    anio = _entero_form(request.args.get('anio'))
    tipo = request.args.get('tipo', '')
    texto = request.args.get('q', '').strip()[:80]
    if anio:
        consulta = consulta.filter(PublicacionCDA.anio == anio)
    if tipo in PUB_TIPOS_SESION:
        consulta = consulta.filter(PublicacionCDA.tipo_sesion == tipo)
    if texto:
        patron = f'%{texto}%'
        consulta = consulta.filter(db.or_(PublicacionCDA.resumen.ilike(patron), PublicacionCDA.palabras_clave.ilike(patron),
                                          PublicacionCDA.titulo.ilike(patron)))
    actas = consulta.all()
    convalidadas = {}
    for c in ConvalidacionCDA.query.filter(ConvalidacionCDA.publicado.is_(True), ConvalidacionCDA.acta_id.isnot(None)).all():
        convalidadas.setdefault(c.acta_id, []).append(c)
    return render_template('actas.html', actas=actas, anios=_anios(PublicacionCDA.anio), tipos=PUB_TIPOS_SESION,
                           filtro={'anio': anio, 'tipo': tipo, 'q': texto}, convalidadas=convalidadas)


@app.route('/transparencia/convalidaciones')
def convalidaciones_publicas():
    consulta = _convalidaciones_publicadas()
    instrumento = request.args.get('instrumento', '')
    estado = request.args.get('estado', '')
    anio = _entero_form(request.args.get('anio'))
    texto = request.args.get('q', '').strip()[:80]
    if instrumento in CONV_INSTRUMENTOS:
        consulta = consulta.filter(ConvalidacionCDA.instrumento == instrumento)
    if estado in CONV_ESTADOS_RUSD:
        consulta = consulta.filter(ConvalidacionCDA.estado_rusd == estado)
    if anio:
        consulta = consulta.filter(db.extract('year', ConvalidacionCDA.fecha_convalidacion) == anio)
    if texto:
        patron = f'%{texto}%'
        consulta = consulta.filter(db.or_(ConvalidacionCDA.entidad.ilike(patron), ConvalidacionCDA.nit.ilike(patron)))
    registros = consulta.all()
    anios = sorted({c.fecha_convalidacion.year for c in _convalidaciones_publicadas().all() if c.fecha_convalidacion}, reverse=True)
    return render_template('convalidaciones.html', registros=registros, instrumentos=CONV_INSTRUMENTOS,
                           estados=CONV_ESTADOS_RUSD, anios=anios,
                           filtro={'instrumento': instrumento, 'estado': estado, 'anio': anio, 'q': texto})


def _documentos_de(categorias):
    docs = (PublicacionCDA.query.filter(PublicacionCDA.categoria.in_(categorias), PublicacionCDA.publicado.is_(True))
            .order_by(PublicacionCDA.fecha.desc().nullslast(), PublicacionCDA.id.desc()).all())
    return {k: [d for d in docs if d.categoria == k] for k in categorias}


@app.route('/transparencia/documentos')
def documentos_publicos():
    return render_template('documentos_cda.html', grupos=_documentos_de(['plan_accion', 'informe_semestral', 'otro']),
                           categorias=PUB_CATEGORIAS, titulo='Planes e informes',
                           descripcion='Plan de acción anual del Consejo, informes semestrales remitidos al Archivo General de la Nación y otros documentos.')


@app.route('/transparencia/reglamento')
def reglamento_publico():
    return render_template('documentos_cda.html', grupos=_documentos_de(['reglamento']),
                           categorias=PUB_CATEGORIAS, titulo='Reglamento interno',
                           descripcion='Reglamento interno de funcionamiento del Consejo Distrital de Archivos de Medellín y actos de creación.')


@app.route('/transparencia/documento/<int:pub_id>')
def documento_publicacion(pub_id):
    pub = PublicacionCDA.query.get_or_404(pub_id)
    if not pub.publicado and not session.get('admin_authenticated'):
        abort(404)
    return _enviar_pdf(pub)


@app.route('/transparencia/convalidaciones/<int:conv_id>/documento')
def documento_convalidacion(conv_id):
    conv = ConvalidacionCDA.query.get_or_404(conv_id)
    if not conv.publicado and not session.get('admin_authenticated'):
        abort(404)
    return _enviar_pdf(conv)


@app.route('/accesibilidad')
def accesibilidad():
    return render_template('politica.html', pagina='accesibilidad')


@app.route('/politica-de-tratamiento-de-datos')
def politica_datos():
    return render_template('politica.html', pagina='datos')


@app.route('/terminos-y-condiciones')
def terminos():
    return render_template('politica.html', pagina='terminos')


@app.route('/mapa-del-sitio')
def mapa_del_sitio():
    return render_template('mapa_sitio.html')


# ---------- Administración ----------

def _formulario_publicacion(pub):
    f = request.form
    categoria = f.get('categoria', 'acta')
    if categoria not in PUB_CATEGORIAS:
        return 'Seleccione una categoría válida.'
    pub.categoria = categoria
    pub.anio = _entero_form(f.get('anio'))
    pub.fecha = _fecha_form(f.get('fecha'))
    pub.resumen = f.get('resumen', '').strip()[:4000]
    pub.palabras_clave = f.get('palabras_clave', '').strip()[:250]
    pub.enlace = _enlace_form(f.get('enlace'))
    pub.publicado = f.get('publicado') == '1'
    if categoria == 'acta':
        pub.numero = _entero_form(f.get('numero'))
        pub.tipo_sesion = f.get('tipo_sesion') if f.get('tipo_sesion') in PUB_TIPOS_SESION else None
        pub.modalidad = f.get('modalidad') if f.get('modalidad') in PUB_MODALIDADES else None
        if not (pub.numero and pub.anio and pub.fecha and pub.tipo_sesion):
            return 'Para un acta indique número, año, fecha de la sesión y tipo de sesión.'
        duplicada = PublicacionCDA.query.filter(PublicacionCDA.categoria == 'acta', PublicacionCDA.numero == pub.numero,
                                                PublicacionCDA.anio == pub.anio, PublicacionCDA.id != (pub.id or 0)).first()
        if duplicada:
            return f'Ya existe el {duplicada.nombre_acta}.'
        pub.titulo = f'Acta N.º {pub.numero:02d} de {pub.anio} · Sesión {pub.tipo_sesion.lower()}'
    else:
        pub.numero = pub.tipo_sesion = pub.modalidad = None
        pub.titulo = f.get('titulo', '').strip()[:250]
        if not pub.titulo:
            return 'Escriba el título del documento.'
    datos, nombre, error = _leer_pdf()
    if error:
        return error
    if datos:
        _asignar_pdf(pub, datos, nombre)
    if not pub.archivo_nombre and not pub.enlace:
        return 'Adjunte el PDF o indique un enlace al documento.'
    if pub.publicado and not pub.fecha_publicacion:
        pub.fecha_publicacion = datetime.utcnow()
    pub.actualizado = datetime.utcnow()
    return None


@app.route('/admin/transparencia', methods=['GET', 'POST'])
@admin_required
def admin_publicaciones():
    pub = None
    if request.method == 'POST':
        pub_id = _entero_form(request.form.get('id'))
        pub = db.session.get(PublicacionCDA, pub_id) if pub_id else PublicacionCDA()
        if pub is None:
            abort(404)
        error = _formulario_publicacion(pub)
        if error:
            flash(error, 'danger')
        else:
            if not pub.id:
                db.session.add(pub)
            db.session.commit()
            flash(f'«{pub.titulo}» guardado{" y publicado" if pub.publicado else " sin publicar"}.', 'success')
            return redirect(url_for('admin_publicaciones'))
    elif request.args.get('editar'):
        pub = db.session.get(PublicacionCDA, _entero_form(request.args.get('editar')) or 0)
    items = PublicacionCDA.query.order_by(PublicacionCDA.categoria, PublicacionCDA.anio.desc().nullslast(),
                                          PublicacionCDA.numero.desc().nullslast(), PublicacionCDA.id.desc()).all()
    return render_template('admin_transparencia.html', items=items, pub=pub, categorias=PUB_CATEGORIAS,
                           tipos=PUB_TIPOS_SESION, modalidades=PUB_MODALIDADES, hoy=datetime.utcnow().date())


@app.route('/admin/transparencia/<int:pub_id>/eliminar', methods=['POST'])
@admin_required
def admin_eliminar_publicacion(pub_id):
    pub = PublicacionCDA.query.get_or_404(pub_id)
    if ConvalidacionCDA.query.filter_by(acta_id=pub.id).count():
        flash('No se puede eliminar: hay convalidaciones asociadas a esta acta. Despublíquela si es necesario.', 'danger')
    else:
        db.session.delete(pub)
        db.session.commit()
        flash('Documento eliminado.', 'success')
    return redirect(url_for('admin_publicaciones'))


def _formulario_convalidacion(conv):
    f = request.form
    conv.entidad = f.get('entidad', '').strip()[:250]
    conv.nit = f.get('nit', '').strip()[:30]
    conv.instrumento = f.get('instrumento') if f.get('instrumento') in CONV_INSTRUMENTOS else None
    conv.tramite = f.get('tramite') if f.get('tramite') in CONV_TRAMITES else None
    conv.version = f.get('version', '').strip()[:40]
    conv.fecha_aprobacion_comite = _fecha_form(f.get('fecha_aprobacion_comite'))
    conv.fecha_convalidacion = _fecha_form(f.get('fecha_convalidacion'))
    acta_id = _entero_form(f.get('acta_id'))
    conv.acta_id = acta_id if acta_id and db.session.get(PublicacionCDA, acta_id) else None
    conv.numero_concepto = f.get('numero_concepto', '').strip()[:60]
    conv.radicado = f.get('radicado', '').strip().upper()[:40]
    conv.estado_rusd = f.get('estado_rusd') if f.get('estado_rusd') in CONV_ESTADOS_RUSD else CONV_ESTADOS_RUSD[0]
    conv.fecha_rusd = _fecha_form(f.get('fecha_rusd'))
    conv.observaciones = f.get('observaciones', '').strip()[:2000]
    conv.enlace = _enlace_form(f.get('enlace'))
    conv.publicado = f.get('publicado') == '1'
    if not (conv.entidad and conv.nit and conv.instrumento and conv.fecha_convalidacion):
        return 'Indique entidad, NIT, instrumento y fecha de convalidación.'
    if not 6 <= len(nit_base(conv.nit)) <= 10:
        return 'Ingrese un NIT válido, por ejemplo 890905211-1.'
    datos, nombre, error = _leer_pdf()
    if error:
        return error
    if datos:
        _asignar_pdf(conv, datos, nombre)
    conv.actualizado = datetime.utcnow()
    return None


@app.route('/admin/convalidaciones', methods=['GET', 'POST'])
@admin_required
def admin_convalidaciones():
    conv = None
    if request.method == 'POST':
        conv_id = _entero_form(request.form.get('id'))
        conv = db.session.get(ConvalidacionCDA, conv_id) if conv_id else ConvalidacionCDA()
        if conv is None:
            abort(404)
        error = _formulario_convalidacion(conv)
        if error:
            flash(error, 'danger')
        else:
            if not conv.id:
                db.session.add(conv)
            db.session.commit()
            flash(f'Convalidación de {conv.instrumento} de {conv.entidad} guardada.', 'success')
            return redirect(url_for('admin_convalidaciones'))
    elif request.args.get('editar'):
        conv = db.session.get(ConvalidacionCDA, _entero_form(request.args.get('editar')) or 0)
    elif request.args.get('radicado'):
        # Precarga desde una solicitud de evaluación de instrumentos (CDA-EV-…).
        s = SolicitudSistema.query.filter_by(radicado=request.args.get('radicado', '').strip().upper()).first()
        if s:
            conv = ConvalidacionCDA(entidad=s.entidad, nit=s.nit, radicado=s.radicado,
                                    instrumento='TVD' if 'TVD' in (s.tema or '') else 'TRD',
                                    tramite='Actualización' if 'actualización' in (s.tema or '').lower() else 'Primera vez',
                                    publicado=True)
    items = ConvalidacionCDA.query.order_by(ConvalidacionCDA.fecha_convalidacion.desc(), ConvalidacionCDA.id.desc()).all()
    actas = PublicacionCDA.query.filter_by(categoria='acta').order_by(PublicacionCDA.anio.desc(), PublicacionCDA.numero.desc()).all()
    return render_template('admin_convalidaciones.html', items=items, conv=conv, actas=actas,
                           instrumentos=CONV_INSTRUMENTOS, tramites=CONV_TRAMITES, estados=CONV_ESTADOS_RUSD)


@app.route('/admin/convalidaciones/<int:conv_id>/eliminar', methods=['POST'])
@admin_required
def admin_eliminar_convalidacion(conv_id):
    db.session.delete(ConvalidacionCDA.query.get_or_404(conv_id))
    db.session.commit()
    flash('Registro de convalidación eliminado.', 'success')
    return redirect(url_for('admin_convalidaciones'))


# ---------------------------------------------------------------------------
# Acompañamiento a entidades priorizadas
# ---------------------------------------------------------------------------
NOMBRE_COMPONENTE = dict(SDA_COMPONENTES)


def _diagnosticos_vigentes():
    """Diagnóstico más reciente de cada entidad (clave -> registro)."""
    vigentes = {}
    for r in Censo.query.order_by(Censo.fecha.desc(), Censo.id.desc()).all():
        vigentes.setdefault(_ficha_clave(r), r)
    return vigentes


def _acompanamientos_por_entidad():
    """Acompañamientos de cada entidad, el más reciente primero (clave -> lista)."""
    por_clave = {}
    for a in AcompanamientoEntidad.query.order_by(AcompanamientoEntidad.fecha_inicio.desc(), AcompanamientoEntidad.id.desc()).all():
        por_clave.setdefault(a.entidad_clave, []).append(a)
    return por_clave


def _estado_acompanamiento(lista):
    if not lista:
        return 'Sin iniciar'
    return 'En curso' if any(a.estado == 'En curso' for a in lista) else 'Cerrado'


def _tipo_acompanamiento(componentes):
    """Integral si 3 o más componentes están en nivel bajo; puntual en los demás casos."""
    en_bajo = sum(1 for v in componentes.values() if v is not None and v * 100 <= UMBRAL_BAJO)
    return 'Integral' if en_bajo >= 3 else 'Puntual'


def _componente_mas_debil(componentes):
    validos = [(v, clave) for clave, _ in SDA_COMPONENTES for v in [componentes.get(clave)] if v is not None]
    return min(validos, key=lambda x: x[0])[1] if validos else SDA_COMPONENTES[0][0]


def _indicadores_acompanamiento():
    """Atención prioritaria y mejora, con el diagnóstico vigente de cada entidad."""
    vigentes = _diagnosticos_vigentes()
    por_clave = _acompanamientos_por_entidad()
    prioritarias = [(k, r) for k, r in vigentes.items() if (r.ima or 0) <= UMBRAL_BAJO]
    atendidas = [k for k, r in prioritarias if any(a.censo_id == r.id or (a.fecha_inicio and r.fecha and a.fecha_inicio >= r.fecha.date()) for a in por_clave.get(k, []))]
    mejoras = []
    for k, lista in por_clave.items():
        cerrado = next((a for a in lista if a.estado == 'Cerrado' and a.fecha_cierre), None)
        actual = vigentes.get(k)
        if cerrado and actual and actual.id != cerrado.censo_id and actual.fecha and actual.fecha.date() >= cerrado.fecha_cierre:
            mejoras.append(round((actual.ima or 0) - (cerrado.nivel_inicial or 0), 1))
    return {
        'prioritarias': len(prioritarias), 'atendidas': len(atendidas),
        'porcentaje': round(len(atendidas) / len(prioritarias) * 100, 1) if prioritarias else None,
        'en_curso': sum(1 for l in por_clave.values() if _estado_acompanamiento(l) == 'En curso'),
        'cerrados': sum(1 for l in por_clave.values() for a in l if a.estado == 'Cerrado'),
        'medidas': len(mejoras), 'mejora_promedio': round(sum(mejoras) / len(mejoras), 1) if mejoras else None,
    }


@app.route('/admin/acompanamiento')
@admin_required
def admin_acompanamiento():
    vigentes = _diagnosticos_vigentes()
    por_clave = _acompanamientos_por_entidad()
    filas = []
    for k, lista in por_clave.items():
        actual = vigentes.get(k)
        for a in lista:
            mejora = None
            if a.estado == 'Cerrado' and a.fecha_cierre and actual and actual.id != a.censo_id and actual.fecha and actual.fecha.date() >= a.fecha_cierre:
                mejora = round((actual.ima or 0) - (a.nivel_inicial or 0), 1)
            filas.append({'a': a, 'componente': NOMBRE_COMPONENTE.get(a.componente, a.componente), 'mejora': mejora,
                          'ficha': actual.id if actual else a.censo_id})
    filas.sort(key=lambda f: (f['a'].estado != 'En curso', -(f['a'].fecha_inicio.toordinal())))
    pendientes = [{'nombre': _sda_nombre(r), 'id': r.id, 'ima': r.ima or 0}
                  for k, r in sorted(vigentes.items(), key=lambda kv: kv[1].ima or 0)
                  if (r.ima or 0) <= UMBRAL_BAJO and _estado_acompanamiento(por_clave.get(k)) == 'Sin iniciar']
    return render_template('admin_acompanamiento.html', filas=filas, pendientes=pendientes,
                           ind=_indicadores_acompanamiento())


@app.route('/admin/diagnostico/<int:censo_id>/acompanamiento', methods=['POST'])
@admin_required
def registrar_acompanamiento(censo_id):
    censo = db.session.get(Censo, censo_id) or abort(404)
    _, actual, componentes = _puntajes_de(censo)
    componente = request.form.get('componente', '')
    responsable = (request.form.get('responsable') or '').strip()
    inicio = _fecha_form(request.form.get('fecha_inicio'))
    if componente not in NOMBRE_COMPONENTE or not responsable or not inicio:
        flash('Complete el componente, el responsable y la fecha de inicio.', 'danger')
        return redirect(url_for('ficha_diagnostico', censo_id=censo_id) + '#acompanamiento')
    db.session.add(AcompanamientoEntidad(
        entidad_clave=_ficha_clave(censo), entidad_nombre=_sda_nombre(censo), censo_id=censo.id,
        nivel_inicial=censo.ima or 0, componente=componente, tipo=_tipo_acompanamiento(componentes),
        responsable=responsable[:200], fecha_inicio=inicio,
        observaciones=(request.form.get('observaciones') or '').strip() or None))
    db.session.commit()
    flash('Acompañamiento registrado.', 'success')
    return redirect(url_for('ficha_diagnostico', censo_id=censo_id) + '#acompanamiento')


@app.route('/admin/acompanamiento/<int:aid>/cerrar', methods=['POST'])
@admin_required
def cerrar_acompanamiento(aid):
    a = db.session.get(AcompanamientoEntidad, aid) or abort(404)
    cierre = _fecha_form(request.form.get('fecha_cierre'))
    destino = request.form.get('volver') or url_for('admin_acompanamiento')
    if not destino.startswith('/') or destino.startswith('//'):
        destino = url_for('admin_acompanamiento')
    if not cierre or cierre < a.fecha_inicio:
        flash('Indique una fecha de cierre igual o posterior a la de inicio.', 'danger')
        return redirect(destino)
    a.estado, a.fecha_cierre = 'Cerrado', cierre
    a.observacion_cierre = (request.form.get('observacion_cierre') or '').strip() or None
    db.session.commit()
    flash('Acompañamiento cerrado. Para medir la mejora, la entidad debe hacer un nuevo diagnóstico.', 'success')
    return redirect(destino)


@app.route('/admin/acompanamiento/<int:aid>/eliminar', methods=['POST'])
@admin_required
def eliminar_acompanamiento(aid):
    a = db.session.get(AcompanamientoEntidad, aid) or abort(404)
    db.session.delete(a)
    db.session.commit()
    flash('Registro de acompañamiento eliminado.', 'success')
    destino = request.form.get('volver') or url_for('admin_acompanamiento')
    return redirect(destino if destino.startswith('/') and not destino.startswith('//') else url_for('admin_acompanamiento'))


@app.errorhandler(404)
def pagina_no_encontrada(_error):
    return render_template('error.html', codigo=404, icono='signpost-split', titulo='No encontramos esta página',
                           mensaje='La dirección no existe, cambió o no tiene permiso para verla. Si llegó desde un enlace del micrositio, escríbanos a consejodarchivo@medellin.gov.co.'), 404


@app.errorhandler(500)
def error_interno(_error):
    db.session.rollback()
    return render_template('error.html', codigo=500, icono='tools', titulo='Algo salió mal',
                           mensaje='Ocurrió un error inesperado. Intente de nuevo en unos minutos; si el problema continúa, escríbanos a consejodarchivo@medellin.gov.co.'), 500


@app.errorhandler(413)
def archivo_muy_grande(_error):
    flash('El archivo supera el tamaño permitido (10 MB).', 'danger')
    return redirect(request.referrer or url_for('admin'))


# ============================================================
# PROTECCIÓN DE LA INFORMACIÓN
# ============================================================

@app.after_request
def cabeceras_seguridad(respuesta):
    respuesta.headers.setdefault('X-Content-Type-Options', 'nosniff')
    respuesta.headers.setdefault('X-Frame-Options', 'SAMEORIGIN')
    respuesta.headers.setdefault('Referrer-Policy', 'strict-origin-when-cross-origin')
    respuesta.headers.setdefault('Permissions-Policy', 'camera=(), microphone=(), geolocation=(), payment=()')
    if app.config.get('SESSION_COOKIE_SECURE'):
        respuesta.headers.setdefault('Strict-Transport-Security', 'max-age=31536000; includeSubDomains')
    if request.path.startswith('/admin') or request.path.startswith('/api/') or request.path.startswith('/resultado'):
        respuesta.headers['Cache-Control'] = 'no-store'
    return respuesta


# ============================================================
# ACCESO DE LAS ENTIDADES CON CÓDIGO POR CORREO
# La entidad escribe su correo institucional y recibe un código de 6 dígitos que
# vence en 10 minutos. Pueden ingresar los correos del directorio de entidades y
# los que el administrador autorice. No se guardan contraseñas.
# ============================================================

import secrets
import smtplib
import ssl
import urllib.request
from email.message import EmailMessage

CODIGO_VIGENCIA_MIN = 10
CODIGO_MAX_INTENTOS = 5
CODIGOS_POR_CORREO_15MIN = 3
SOLICITUDES_POR_IP_15MIN = 10
ENTIDAD_INACTIVIDAD_SEG = 2 * 60 * 60
_SOLICITUDES_IP = {}


class AccesoEntidad(db.Model):
    """Correos autorizados para ingresar en nombre de una entidad."""
    __tablename__ = 'acceso_entidad'
    id = db.Column(db.Integer, primary_key=True)
    correo = db.Column(db.String(200), nullable=False, index=True)
    entidad_clave = db.Column(db.String(120), nullable=False, index=True)
    entidad_nombre = db.Column(db.String(250), nullable=False)
    origen = db.Column(db.String(20), default='manual')  # directorio | manual
    activo = db.Column(db.Boolean, default=True, nullable=False)
    ultimo_ingreso = db.Column(db.DateTime)
    ingresos = db.Column(db.Integer, default=0)
    creado = db.Column(db.DateTime, default=datetime.utcnow)


class EntidadAcceso(db.Model):
    """NIT e ingresos de cada entidad del directorio. La clave son los últimos 4 dígitos del NIT."""
    __tablename__ = 'entidad_acceso'
    entidad_clave = db.Column(db.String(120), primary_key=True)
    nit = db.Column(db.String(30))
    clave_hash = db.Column(db.String(64))
    clave_sal = db.Column(db.String(32))
    intentos_fallidos = db.Column(db.Integer, default=0)
    bloqueado_hasta = db.Column(db.DateTime)
    ultimo_ingreso = db.Column(db.DateTime)
    ingresos = db.Column(db.Integer, default=0)
    actualizado = db.Column(db.DateTime, default=datetime.utcnow)


INTENTOS_NIT_MAX = 5
BLOQUEO_NIT_MIN = 15


def _ingreso_por_nit(usuario, clave):
    """Devuelve (lista de (clave_entidad, nombre), mensaje_error)."""
    base = nit_base(usuario)
    if not 6 <= len(base) <= 10:
        return [], 'Escriba el NIT sin dígito de verificación, por ejemplo 890905211.'
    candidatas = [(clave_de_fila(f), f.get('Entidad', '')) for f in obtener_directorio() if nit_base(f.get('NIT')) == base]
    validas, bloqueada = [], False
    for clave_ent, nombre in candidatas:
        acc = db.session.get(EntidadAcceso, clave_ent)
        if acc and acc.bloqueado_hasta and acc.bloqueado_hasta > datetime.utcnow():
            bloqueada = True
            continue
        correcta = hmac.compare_digest(clave.strip(), base[-4:])
        if correcta:
            validas.append((clave_ent, nombre))
        else:
            if acc is None:
                acc = EntidadAcceso(entidad_clave=clave_ent)
                db.session.add(acc)
            acc.intentos_fallidos = (acc.intentos_fallidos or 0) + 1
            if acc.intentos_fallidos >= INTENTOS_NIT_MAX:
                acc.bloqueado_hasta = datetime.utcnow() + timedelta(minutes=BLOQUEO_NIT_MIN)
                acc.intentos_fallidos = 0
                app.logger.warning('Ingreso por NIT bloqueado para %s', clave_ent)
    db.session.commit()
    if validas:
        return validas, None
    if bloqueada:
        return [], f'Se superaron los intentos permitidos. Espere {BLOQUEO_NIT_MIN} minutos e intente de nuevo.'
    return [], 'El NIT o la clave no son correctos.'


def _registrar_ingreso_nit(clave_ent, nombre, usuario):
    acc = db.session.get(EntidadAcceso, clave_ent)
    if acc is None:
        acc = EntidadAcceso(entidad_clave=clave_ent)
        db.session.add(acc)
    acc.intentos_fallidos = 0
    acc.bloqueado_hasta = None
    acc.ultimo_ingreso = datetime.utcnow()
    acc.ingresos = (acc.ingresos or 0) + 1
    db.session.commit()
    session.pop('acceso_pendiente', None)
    session['entidad'] = {'clave': clave_ent, 'nombre': nombre, 'correo': '', 'usuario': nit_base(usuario),
                          't': time.time()}
    app.logger.info('Ingreso por NIT de la entidad %s', clave_ent)


class CodigoAcceso(db.Model):
    __tablename__ = 'codigo_acceso'
    id = db.Column(db.Integer, primary_key=True)
    correo = db.Column(db.String(200), nullable=False, index=True)
    codigo_hash = db.Column(db.String(64), nullable=False)
    expira = db.Column(db.DateTime, nullable=False)
    intentos = db.Column(db.Integer, default=0)
    usado = db.Column(db.Boolean, default=False)
    ip = db.Column(db.String(64))
    creado = db.Column(db.DateTime, default=datetime.utcnow)


def _correos_en(texto):
    return [c.lower() for c in re.findall(r'[\w.+-]+@[\w-]+(?:\.[\w-]+)+', texto or '')]


def clave_de_fila(fila):
    codigo = str(fila.get('Código', '')).strip()
    return codigo or 'N:' + clave_entidad(fila.get('Entidad', ''))[:110]


def fila_por_clave(clave):
    return next((f for f in obtener_directorio() if clave_de_fila(f) == clave), None)


def entidades_para_correo(correo):
    """Entidades que puede representar un correo: del directorio y autorizaciones manuales."""
    correo = (correo or '').strip().lower()
    if not correo:
        return []
    bloqueadas = {a.entidad_clave for a in AccesoEntidad.query.filter_by(correo=correo, activo=False).all()}
    resultado = {}
    for fila in obtener_directorio():
        if correo in _correos_en(fila.get('Correo Responsable GD', '')):
            clave = clave_de_fila(fila)
            if clave not in bloqueadas:
                resultado[clave] = fila.get('Entidad', '')
    for a in AccesoEntidad.query.filter_by(correo=correo, activo=True, origen='manual').all():
        resultado[a.entidad_clave] = a.entidad_nombre
    return sorted(resultado.items(), key=lambda kv: kv[1])


def correo_configurado():
    if os.environ.get('BREVO_API_KEY'):
        return 'Brevo (API)'
    if os.environ.get('SMTP_HOST'):
        return 'SMTP ' + os.environ.get('SMTP_HOST')
    return None


def modo_acceso():
    """Variable ACCESO_ENTIDADES en Render:
    - vacía o «nit» (predeterminado): usuario = NIT de la entidad y clave = últimos 4 dígitos
      del NIT.
    - «correo»: código de un solo uso enviado por correo (requiere servicio de correo).
    - «prueba»: como «correo», pero el código se muestra en pantalla.
    - «no»: el diagnóstico queda abierto, sin ingreso."""
    ajuste = os.environ.get('ACCESO_ENTIDADES', '').strip().lower()
    if ajuste in ('no', '0', 'false'):
        return None
    if ajuste == 'prueba':
        return 'prueba'
    if ajuste in ('correo', 'si', 'sí'):
        return 'correo'
    return 'nit'


def acceso_requerido():
    return modo_acceso() is not None


def enviar_correo(destino, asunto, texto, html):
    remitente = os.environ.get('MAIL_FROM', 'consejodarchivo@medellin.gov.co')
    nombre = os.environ.get('MAIL_FROM_NAME', 'Consejo Distrital de Archivos de Medellín')
    if os.environ.get('BREVO_API_KEY'):
        datos = json.dumps({'sender': {'email': remitente, 'name': nombre}, 'to': [{'email': destino}],
                            'subject': asunto, 'textContent': texto, 'htmlContent': html}).encode()
        req = urllib.request.Request('https://api.brevo.com/v3/smtp/email', data=datos, method='POST',
                                     headers={'api-key': os.environ['BREVO_API_KEY'], 'Content-Type': 'application/json',
                                              'Accept': 'application/json'})
        with urllib.request.urlopen(req, timeout=15) as r:
            return 200 <= r.status < 300
    if os.environ.get('SMTP_HOST'):
        msg = EmailMessage()
        msg['From'] = f'{nombre} <{remitente}>'
        msg['To'] = destino
        msg['Subject'] = asunto
        msg.set_content(texto)
        msg.add_alternative(html, subtype='html')
        puerto = int(os.environ.get('SMTP_PORT', '587'))
        contexto = ssl.create_default_context()
        if puerto == 465:
            servidor = smtplib.SMTP_SSL(os.environ['SMTP_HOST'], puerto, timeout=15, context=contexto)
        else:
            servidor = smtplib.SMTP(os.environ['SMTP_HOST'], puerto, timeout=15)
            servidor.starttls(context=contexto)
        with servidor:
            if os.environ.get('SMTP_USER'):
                servidor.login(os.environ['SMTP_USER'], os.environ.get('SMTP_PASSWORD', ''))
            servidor.send_message(msg)
        return True
    # Sin servicio de correo (solo pruebas locales): el código queda en el registro.
    app.logger.warning('Código de acceso para %s (modo de prueba, sin correo configurado): %s', destino, texto)
    return True


def _hash_codigo(correo, codigo):
    return hmac.new(app.secret_key.encode() if isinstance(app.secret_key, str) else app.secret_key,
                    f'{correo}:{codigo}'.encode(), hashlib.sha256).hexdigest()


def entidad_en_sesion():
    e = session.get('entidad')
    if not e:
        return None
    if time.time() - e.get('t', 0) > ENTIDAD_INACTIVIDAD_SEG:
        session.pop('entidad', None)
        return None
    e['t'] = time.time()
    session['entidad'] = e
    return e


@app.context_processor
def inyectar_entidad():
    return {'entidad_sesion': session.get('entidad'), 'modo_acceso_entidades': modo_acceso()}


def _iniciar_sesion_entidad(correo, clave, nombre):
    acceso = AccesoEntidad.query.filter_by(correo=correo, entidad_clave=clave).first()
    if acceso is None:
        acceso = AccesoEntidad(correo=correo, entidad_clave=clave, entidad_nombre=nombre, origen='directorio')
        db.session.add(acceso)
    acceso.ultimo_ingreso = datetime.utcnow()
    acceso.ingresos = (acceso.ingresos or 0) + 1
    db.session.commit()
    session.pop('acceso_pendiente', None)
    session['entidad'] = {'clave': clave, 'nombre': nombre, 'correo': correo, 't': time.time()}
    app.logger.info('Ingreso de entidad %s con %s', clave, correo)


@app.route('/acceso', methods=['GET', 'POST'])
def acceso_entidad():
    siguiente = request.values.get('next') or url_for('censo')
    if not siguiente.startswith('/') or siguiente.startswith('//'):
        siguiente = url_for('censo')
    if modo_acceso() == 'nit':
        return _acceso_por_nit(siguiente)
    paso = 'correo'
    correo = (request.form.get('correo') or session.get('acceso_correo') or '').strip().lower()
    if request.method == 'POST' and request.form.get('accion') == 'enviar':
        ip = _ip_cliente()
        recientes = [t for t in _SOLICITUDES_IP.get(ip, []) if time.time() - t < 900]
        if len(recientes) >= SOLICITUDES_POR_IP_15MIN:
            flash('Se hicieron demasiadas solicitudes desde esta conexión. Espere unos minutos.', 'danger')
            return render_template('acceso.html', paso='correo', correo=correo, siguiente=siguiente), 429
        _SOLICITUDES_IP[ip] = recientes + [time.time()]
        if not re.fullmatch(r'[\w.+-]+@[\w-]+(?:\.[\w-]+)+', correo):
            flash('Escriba un correo electrónico válido.', 'danger')
            return render_template('acceso.html', paso='correo', correo=correo, siguiente=siguiente)
        hace15 = datetime.utcnow() - timedelta(minutes=15)
        if CodigoAcceso.query.filter(CodigoAcceso.correo == correo, CodigoAcceso.creado > hace15).count() >= CODIGOS_POR_CORREO_15MIN:
            flash('Ya se enviaron varios códigos a este correo. Revise su bandeja o espere 15 minutos.', 'warning')
        elif entidades_para_correo(correo):
            codigo = f'{secrets.randbelow(1000000):06d}'
            db.session.add(CodigoAcceso(correo=correo, codigo_hash=_hash_codigo(correo, codigo), ip=ip,
                                        expira=datetime.utcnow() + timedelta(minutes=CODIGO_VIGENCIA_MIN)))
            db.session.commit()
            texto = (f'Su código de acceso al Diagnóstico Territorial Archivístico es {codigo}.\n\n'
                     f'Vence en {CODIGO_VIGENCIA_MIN} minutos. Si usted no lo solicitó, ignore este mensaje.\n\n'
                     'Consejo Distrital de Archivos de Medellín')
            html = (f'<p>Su código de acceso al <strong>Diagnóstico Territorial Archivístico</strong> es:</p>'
                    f'<p style="font-size:28px;font-weight:bold;letter-spacing:6px">{codigo}</p>'
                    f'<p>Vence en {CODIGO_VIGENCIA_MIN} minutos. Si usted no lo solicitó, ignore este mensaje.</p>'
                    '<p>Consejo Distrital de Archivos de Medellín</p>')
            if modo_acceso() == 'prueba':
                flash(f'MODO DE PRUEBA · Su código es {codigo}. En producción este código llega por correo.', 'warning')
            else:
              try:
                enviar_correo(correo, f'Código de acceso: {codigo}', texto, html)
              except Exception:
                app.logger.exception('No se pudo enviar el código de acceso')
                flash('No fue posible enviar el correo en este momento. Intente de nuevo más tarde.', 'danger')
                return render_template('acceso.html', paso='correo', correo=correo, siguiente=siguiente)
        else:
            app.logger.info('Solicitud de código para un correo no autorizado: %s', correo)
        # La respuesta es la misma exista o no el correo, para no revelar quién está registrado.
        session['acceso_correo'] = correo
        flash(f'Si el correo está autorizado para una entidad, en unos segundos recibirá un código de 6 dígitos. Revise también la carpeta de correo no deseado.', 'info')
        paso = 'codigo'
    elif request.method == 'POST' and request.form.get('accion') == 'verificar':
        paso = 'codigo'
        digitado = re.sub(r'\D', '', request.form.get('codigo', ''))
        registro = (CodigoAcceso.query.filter(CodigoAcceso.correo == correo, CodigoAcceso.usado.is_(False),
                                              CodigoAcceso.expira > datetime.utcnow())
                    .order_by(CodigoAcceso.id.desc()).first())
        if registro is None:
            flash('El código venció o no es válido. Solicite uno nuevo.', 'danger')
            paso = 'correo'
        elif registro.intentos >= CODIGO_MAX_INTENTOS:
            registro.usado = True
            db.session.commit()
            flash('Se superó el número de intentos. Solicite un código nuevo.', 'danger')
            paso = 'correo'
        elif not hmac.compare_digest(registro.codigo_hash, _hash_codigo(correo, digitado)):
            registro.intentos += 1
            db.session.commit()
            flash(f'El código no es correcto. Le quedan {CODIGO_MAX_INTENTOS - registro.intentos} intentos.', 'danger')
        else:
            registro.usado = True
            db.session.commit()
            entidades = entidades_para_correo(correo)
            if len(entidades) == 1:
                _iniciar_sesion_entidad(correo, *entidades[0])
                flash(f'Ingresó como {entidades[0][1]}.', 'success')
                return redirect(siguiente)
            if not entidades:
                flash('Este correo ya no está autorizado para ninguna entidad.', 'danger')
                paso = 'correo'
            else:
                session['acceso_pendiente'] = {'correo': correo, 'claves': [c for c, _ in entidades], 't': time.time()}
                return render_template('acceso.html', paso='entidad', entidades=entidades, siguiente=siguiente)
    elif request.method == 'POST' and request.form.get('accion') == 'elegir':
        pendiente = session.get('acceso_pendiente') or {}
        clave = request.form.get('entidad', '')
        if pendiente and time.time() - pendiente.get('t', 0) < 600 and clave in pendiente.get('claves', []):
            nombre = dict(entidades_para_correo(pendiente['correo'])).get(clave)
            if nombre:
                _iniciar_sesion_entidad(pendiente['correo'], clave, nombre)
                flash(f'Ingresó como {nombre}.', 'success')
                return redirect(siguiente)
        flash('La selección venció. Solicite un código nuevo.', 'danger')
    elif request.args.get('nuevo'):
        session.pop('acceso_correo', None)
        correo = ''
    return render_template('acceso.html', paso=paso, correo=correo, siguiente=siguiente)


def _acceso_por_nit(siguiente):
    usuario = request.form.get('usuario', '').strip()
    if request.method == 'POST' and request.form.get('accion') == 'elegir':
        pendiente = session.get('acceso_pendiente') or {}
        clave_ent = request.form.get('entidad', '')
        if pendiente and time.time() - pendiente.get('t', 0) < 600 and clave_ent in dict(pendiente.get('opciones', [])):
            nombre = dict(pendiente['opciones'])[clave_ent]
            _registrar_ingreso_nit(clave_ent, nombre, pendiente['usuario'])
            flash(f'Ingresó como {nombre}.', 'success')
            return redirect(siguiente)
        flash('La selección venció. Ingrese de nuevo.', 'danger')
    elif request.method == 'POST':
        ip = _ip_cliente()
        recientes = [t for t in _SOLICITUDES_IP.get(ip, []) if time.time() - t < 900]
        if len(recientes) >= 20:
            flash('Demasiados intentos desde esta conexión. Espere 15 minutos.', 'danger')
            return render_template('acceso.html', paso='nit', usuario=usuario, siguiente=siguiente), 429
        _SOLICITUDES_IP[ip] = recientes + [time.time()]
        validas, error = _ingreso_por_nit(usuario, request.form.get('clave', ''))
        if error:
            flash(error, 'danger')
        elif len(validas) == 1:
            _registrar_ingreso_nit(validas[0][0], validas[0][1], usuario)
            _SOLICITUDES_IP.pop(ip, None)
            flash(f'Ingresó como {validas[0][1]}.', 'success')
            return redirect(siguiente)
        else:
            session['acceso_pendiente'] = {'usuario': nit_base(usuario), 'opciones': validas, 't': time.time()}
            return render_template('acceso.html', paso='entidad', entidades=validas, siguiente=siguiente)
    return render_template('acceso.html', paso='nit', usuario=usuario, siguiente=siguiente)


@app.route('/acceso/salir', methods=['POST'])
def salir_entidad():
    session.pop('entidad', None)
    session.pop('acceso_correo', None)
    flash('Cerró la sesión de su entidad.', 'success')
    return redirect(url_for('index'))


def _diagnosticos_de(entidad):
    fila = fila_por_clave(entidad['clave'])
    codigo = str(fila.get('Código', '')).strip() if fila else ''
    nombre_clave = clave_entidad(entidad['nombre'])
    return [c for c in Censo.query.order_by(Censo.fecha.desc()).all()
            if (codigo and c.codigo_entidad == codigo) or clave_entidad((c.nombre_entidad or '').split('|')[0]) == nombre_clave]


@app.route('/mi-entidad')
def mi_entidad():
    entidad = entidad_en_sesion()
    if not entidad:
        return redirect(url_for('acceso_entidad', next=request.path))
    solicitudes = [s for s in SolicitudSistema.query.order_by(SolicitudSistema.id.desc()).all()
                   if clave_entidad(s.entidad or '') == clave_entidad(entidad['nombre'])]
    return render_template('mi_entidad.html', entidad=entidad, diagnosticos=_diagnosticos_de(entidad),
                           solicitudes=solicitudes)


@app.route('/admin/accesos', methods=['GET', 'POST'])
@admin_required
def admin_accesos():
    if request.method == 'POST':
        accion = request.form.get('accion')
        if accion == 'autorizar':
            correo = request.form.get('correo', '').strip().lower()
            fila = fila_por_clave(request.form.get('entidad', ''))
            if not fila or not re.fullmatch(r'[\w.+-]+@[\w-]+(?:\.[\w-]+)+', correo):
                flash('Seleccione la entidad y escriba un correo válido.', 'danger')
            else:
                clave = clave_de_fila(fila)
                a = AccesoEntidad.query.filter_by(correo=correo, entidad_clave=clave).first() or \
                    AccesoEntidad(correo=correo, entidad_clave=clave, entidad_nombre=fila.get('Entidad', ''))
                a.origen, a.activo = 'manual', True
                db.session.add(a)
                db.session.commit()
                flash(f'{correo} quedó autorizado para {a.entidad_nombre}.', 'success')
        elif accion in ('bloquear', 'activar'):
            correo = request.form.get('correo', '').strip().lower()
            clave = request.form.get('entidad', '')
            fila = fila_por_clave(clave)
            a = AccesoEntidad.query.filter_by(correo=correo, entidad_clave=clave).first()
            if a is None and fila:
                a = AccesoEntidad(correo=correo, entidad_clave=clave, entidad_nombre=fila.get('Entidad', ''), origen='directorio')
                db.session.add(a)
            if a:
                a.activo = accion == 'activar'
                db.session.commit()
                flash(f'Acceso de {correo} {"activado" if a.activo else "bloqueado"}.', 'success')
        elif accion == 'nit':
            clave_ent = request.form.get('entidad', '')
            if fila_por_clave(clave_ent) is None:
                abort(404)
            nit = request.form.get('nit', '').strip()[:30]
            if nit and not 6 <= len(nit_base(nit)) <= 10:
                flash('Ingrese un NIT válido, por ejemplo 890905211-1.', 'danger')
            else:
                acc = db.session.get(EntidadAcceso, clave_ent) or EntidadAcceso(entidad_clave=clave_ent)
                acc.nit = nit or None
                acc.actualizado = datetime.utcnow()
                db.session.add(acc)
                db.session.commit()
                flash('NIT guardado.' if nit else 'NIT eliminado.', 'success')
        elif accion == 'restablecer':
            acc = db.session.get(EntidadAcceso, request.form.get('entidad', ''))
            if acc:
                acc.intentos_fallidos, acc.bloqueado_hasta = 0, None
                db.session.commit()
                flash('Se quitó el bloqueo de ingreso.', 'success')
        elif accion == 'prueba':
            correo = request.form.get('correo', '').strip()
            try:
                enviar_correo(correo, 'Prueba de correo del micrositio', 'El servicio de correo del micrositio funciona.',
                              '<p>El servicio de correo del micrositio funciona.</p>')
                flash(f'Correo de prueba enviado a {correo}.', 'success')
            except Exception as error:
                app.logger.exception('Falló el correo de prueba')
                flash(f'No se pudo enviar el correo de prueba: {error}', 'danger')
        return redirect(url_for('admin_accesos'))
    registros = {(a.correo, a.entidad_clave): a for a in AccesoEntidad.query.all()}
    filas = []
    for fila in obtener_directorio():
        clave = clave_de_fila(fila)
        correos = [(c, 'directorio') for c in dict.fromkeys(_correos_en(fila.get('Correo Responsable GD', '')))]
        correos += [(a.correo, 'manual') for a in registros.values() if a.entidad_clave == clave and a.origen == 'manual']
        acc = db.session.get(EntidadAcceso, clave)
        filas.append({'clave': clave, 'nombre': fila.get('Entidad', ''), 'sigla': fila.get('Sigla', ''),
                      'nit': fila.get('NIT', ''), 'acceso': acc,
                      'correos': [{'correo': c, 'origen': o, 'registro': registros.get((c, clave))} for c, o in correos]})
    filas.sort(key=lambda f: (bool(f['nit']) if modo_acceso() == 'nit' else bool(f['correos']), f['nombre']))
    return render_template('admin_accesos.html', filas=filas, servicio=correo_configurado(), modo=modo_acceso(), ahora=datetime.utcnow())


def _sql_default(columna, dialecto):
    """Valor DEFAULT en SQL para una columna nueva, o None si no hay uno fijo."""
    default = columna.default
    if default is None:
        return None
    if getattr(default, 'is_scalar', False):
        valor = default.arg
        if isinstance(valor, bool):
            return 'TRUE' if valor else 'FALSE'
        if isinstance(valor, (int, float)):
            return str(valor)
        if isinstance(valor, str):
            return "'" + valor.replace("'", "''") + "'"
        return None
    # Fechas con default de Python (datetime.utcnow / datetime.now).
    if isinstance(columna.type, db.DateTime):
        return 'CURRENT_TIMESTAMP'
    return None


def sincronizar_esquema():
    """Agrega a cada tabla existente las columnas del modelo que le falten.

    - Columna con valor por defecto fijo: se agrega con ese DEFAULT (y NOT NULL si el
      modelo lo exige); las filas existentes reciben el valor por defecto.
    - Columna sin valor por defecto: se agrega aceptando vacíos, para no fallar con
      filas antiguas.
    - En PostgreSQL, las columnas antiguas que ya no usa el modelo y exigen valor
      (NOT NULL sin DEFAULT) pasan a aceptar vacíos, para que los INSERT nuevos no fallen.
    Nunca se borran tablas, columnas ni registros.
    """
    try:
        inspector = sqlalchemy_inspect(db.engine)
        dialecto = db.engine.dialect
        tablas_existentes = set(inspector.get_table_names())
        for tabla in db.metadata.sorted_tables:
            if tabla.name not in tablas_existentes:
                continue
            columnas_bd = {c['name']: c for c in inspector.get_columns(tabla.name)}
            for columna in tabla.columns:
                if columna.name in columnas_bd:
                    continue
                tipo = columna.type.compile(dialect=dialecto)
                default = _sql_default(columna, dialecto)
                sql = f'ALTER TABLE {tabla.name} ADD COLUMN {columna.name} {tipo}'
                relleno = None
                if default == 'CURRENT_TIMESTAMP':
                    # SQLite no admite un DEFAULT calculado al agregar columnas: se agrega
                    # la columna y se llenan las filas existentes con la fecha actual.
                    # Las filas nuevas reciben la fecha desde el modelo (Python).
                    relleno = f'UPDATE {tabla.name} SET {columna.name} = CURRENT_TIMESTAMP WHERE {columna.name} IS NULL'
                elif default is not None:
                    sql += f' DEFAULT {default}'
                    if not columna.nullable and not columna.primary_key:
                        sql += ' NOT NULL'
                try:
                    with db.engine.begin() as conexion:
                        conexion.exec_driver_sql(sql)
                        if relleno:
                            conexion.exec_driver_sql(relleno)
                    app.logger.warning('Esquema actualizado: %s', sql)
                except Exception:
                    app.logger.exception('No se pudo agregar la columna %s.%s', tabla.name, columna.name)
            if dialecto.name == 'postgresql':
                del_modelo = {c.name for c in tabla.columns}
                for nombre, info in columnas_bd.items():
                    if nombre in del_modelo or info.get('nullable', True) or info.get('default') is not None:
                        continue
                    sql = f'ALTER TABLE {tabla.name} ALTER COLUMN {nombre} DROP NOT NULL'
                    try:
                        with db.engine.begin() as conexion:
                            conexion.exec_driver_sql(sql)
                        app.logger.warning('Esquema actualizado: %s', sql)
                    except Exception:
                        app.logger.exception('No se pudo flexibilizar la columna %s.%s', tabla.name, nombre)
    except Exception:
        app.logger.exception('No se pudo sincronizar el esquema de la base de datos')


with app.app_context():
    db.create_all()
    # Migración compatible: create_all() crea tablas nuevas, pero no agrega columnas a
    # tablas que ya existían (por ejemplo, una tabla creada por una versión anterior en
    # Neon). Aquí se compara cada tabla con su modelo y se agregan las columnas que
    # falten, sin borrar ni modificar datos. Ver sincronizar_esquema().
    sincronizar_esquema()

    # Recalcula los diagnósticos guardados con la escala vigente (0 – 0,5 – 1), para
    # que el índice guardado coincida con el de la ficha. Solo cambia puntaje_obtenido,
    # variables_aplicables, ima y nivel; las respuestas no se tocan. Aplica al cuestionario
    # V04 (si se ajusta cuestionario.py, los resultados se actualizan) y al formulario de
    # 10 preguntas. Los registros más antiguos (sin q1_politica) se dejan como están.
    try:
        actualizados = 0
        for registro in Censo.query.all():
            try:
                datos = json.loads(registro.respuestas_json or '{}')
            except (TypeError, ValueError):
                continue
            if not isinstance(datos, dict) or not (cq.es_v04(datos) or datos.get('q1_politica')):
                continue
            total, aplicables, ima, nivel = calcular_diagnostico_nuevo(datos)
            if (registro.ima, registro.nivel, registro.puntaje_obtenido, registro.variables_aplicables) != (ima, nivel, total, aplicables):
                registro.puntaje_obtenido, registro.variables_aplicables = total, aplicables
                registro.ima, registro.nivel = ima, nivel
                actualizados += 1
        if actualizados:
            db.session.commit()
            app.logger.info('Diagnósticos recalculados con la escala vigente: %s', actualizados)
    except Exception:
        db.session.rollback()
        app.logger.exception('No se pudieron recalcular los diagnósticos con la escala vigente')

    # Calcula la huella SHA-256 de los PDF que se cargaron sin ella, para que todos los
    # documentos publicados se puedan verificar. Solo completa el campo vacío.
    try:
        completados = 0
        for modelo in (PublicacionCDA, ConvalidacionCDA):
            for obj in modelo.query.filter(modelo.archivo_nombre.isnot(None), modelo.archivo_hash.is_(None)).all():
                if obj.archivo:
                    obj.archivo_hash = hashlib.sha256(obj.archivo).hexdigest()
                    completados += 1
        if completados:
            db.session.commit()
            app.logger.info('Huellas SHA-256 completadas: %s', completados)
    except Exception:
        db.session.rollback()
        app.logger.exception('No se pudieron completar las huellas SHA-256 de los PDF')


if __name__ == "__main__":
    app.run(debug=os.environ.get("FLASK_DEBUG", "false").lower() == "true")
