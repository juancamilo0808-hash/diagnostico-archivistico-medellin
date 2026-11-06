from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file, jsonify, abort
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import inspect as sqlalchemy_inspect
from datetime import datetime, timedelta
from pathlib import Path
from functools import wraps
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import json
import math
import os
import csv
import unicodedata
import re

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

NORMATIVIDAD=[
    {
        "categoria":"Ley",
        "titulo":"Ley 594 de 2000",
        "descripcion":"Ley General de Archivos y marco general de la función archivística del Estado.",
        "estado":"Vigente",
        "icono":"bi-journal-text",
        "web":"https://normativa.archivogeneral.gov.co/ley-594-de-2000/",
        "pdf":"https://normativa.archivogeneral.gov.co/ley-594-de-2000/?pdf=41"
    },
    {
        "categoria":"Ley",
        "titulo":"Ley 1712 de 2014",
        "descripcion":"Transparencia y derecho de acceso a la información pública nacional.",
        "estado":"Vigente",
        "icono":"bi-shield-check",
        "web":"https://normativa.archivogeneral.gov.co/ley-1712-de-2014/",
        "pdf":"https://normativa.archivogeneral.gov.co/ley-1712-de-2014/?pdf=418"
    },
    {
        "categoria":"Decreto",
        "titulo":"Decreto 1080 de 2015",
        "descripcion":"Decreto Único Reglamentario del Sector Cultura, con disposiciones aplicables a la gestión documental.",
        "estado":"Vigente",
        "icono":"bi-file-earmark-ruled",
        "web":"https://normativa.archivogeneral.gov.co/decreto-1080-de-2015/",
        "pdf":"https://normativa.archivogeneral.gov.co/decreto-1080-de-2015/?pdf=33"
    },
    {
        "categoria":"Acuerdo",
        "titulo":"Acuerdo No. 001 de 2024",
        "descripcion":"Acuerdo Único de la Función Archivística y criterios técnicos y jurídicos para su implementación.",
        "estado":"Vigente",
        "icono":"bi-file-earmark-medical",
        "web":"https://normativa.archivogeneral.gov.co/acuerdo-no-001-del-2024/",
        "pdf":"https://normativa.archivogeneral.gov.co/wp-content/uploads/2024/04/2024-02_29_AcuerdoAGN-FIRMADO.pdf"
    },
    {
        "categoria":"Resolución",
        "titulo":"Resolución MinTIC 1519 de 2020",
        "descripcion":"Lineamientos de publicación, accesibilidad web, transparencia, seguridad digital y datos abiertos.",
        "estado":"Fuente normativa oficial",
        "icono":"bi-gear",
        "web":"https://normograma.mintic.gov.co/mintic/compilacion/docs/resolucion_mintic_1519_2020.htm",
        "pdf":""
    },
    {
        "categoria":"Circular",
        "titulo":"Circulares y directrices archivísticas del AGN",
        "descripcion":"Consulta las circulares y orientaciones oficiales emitidas para la función archivística.",
        "estado":"Repositorio oficial de circulares",
        "icono":"bi-file-earmark",
        "web":"https://normativa.archivogeneral.gov.co/inicio/circulares/",
        "pdf":""
    },
    {
        "categoria":"Guía y manual",
        "titulo":"Guías y manuales de gestión documental",
        "descripcion":"Material técnico de orientación para organizar, conservar y administrar documentos y archivos.",
        "estado":"Repositorio oficial de publicaciones",
        "icono":"bi-book",
        "web":"https://www.archivogeneral.gov.co/consulte/recursos/publicaciones",
        "pdf":""
    },
    {
        "categoria":"Instrumento archivístico",
        "titulo":"Instrumentos archivísticos: TRD, TVD, PGD y PINAR",
        "descripcion":"Referentes y recursos para la elaboración e implementación de instrumentos archivísticos.",
        "estado":"Portal oficial de instrumentos",
        "icono":"bi-stack",
        "web":"https://www.archivogeneral.gov.co/instrumentos-archivisticos",
        "pdf":""
    },
    {
        "categoria":"Documento formativo",
        "titulo":"Material formativo en gestión documental y archivos",
        "descripcion":"Recursos de capacitación, cartillas, infografías y materiales pedagógicos para fortalecer conocimientos archivísticos.",
        "estado":"Repositorio oficial de publicaciones",
        "icono":"bi-mortarboard",
        "web":"https://www.archivogeneral.gov.co/consulte/recursos/publicaciones",
        "pdf":""
    },
    {
        "categoria":"Referente técnico",
        "titulo":"MGDA – Modelo de Gestión Documental y Administración de Archivos",
        "descripcion":"Referente técnico del AGN para diagnóstico, gestión y desarrollo de la función archivística. La copia oficial disponible corresponde a la versión 2.0 (2020).",
        "estado":"Documento técnico",
        "icono":"bi-diagram-3",
        "web":"https://www.archivogeneral.gov.co/",
        "pdf":"https://www.archivogeneral.gov.co/sites/default/files/Estructura_Web/5_Consulte/Recursos/Publicacionees/V3_MGDA-min.pdf"
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
    if ima <= 33.33: nivel = 'Bajo'
    elif ima <= 66.66: nivel = 'Medio'
    else: nivel = 'Alto'
    return total,aplicables,ima,nivel,q6

@app.route('/')
def index(): return render_template('index.html')



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


def obtener_directorio():
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


@app.route('/api/buscar-entidades')
def api_buscar_entidades():
    """Devuelve coincidencias para el buscador del formulario."""
    termino = request.args.get('q', '').strip()

    if len(termino) < 2:
        return jsonify([])

    resultados = buscar_entidades_directorio(termino)

    return jsonify([
        {
            'nombre': resultado['nombre'],
            'sigla': resultado['sigla'],
            'datos': resultado['datos']
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


@app.route('/api/borrador/<codigo>', methods=['GET'])
def api_obtener_borrador(codigo):
    codigo = validar_codigo_borrador(codigo)
    if not codigo:
        return jsonify({'ok': False, 'mensaje': 'El código de recuperación no es válido.'}), 400
    try:
        borrador = db.session.get(BorradorDiagnostico, codigo)
        if borrador is None:
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
        if borrador is None:
            borrador = BorradorDiagnostico(id=codigo, datos_json=serializado, estado='borrador')
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


@app.route('/diagnostico', methods=['GET', 'POST'])
def censo():
    if request.method == 'POST':
        obligatorios = ['nombre_entidad', 'nit_entidad', 'naturaleza', 'sector',
                        'responsable_nombre', 'responsable_cargo', 'responsable_correo',
                        'q1_politica', 'q2_seguimiento', 'q3_personal', 'q4_depositos',
                        'q5_trd', 'q6_procesos', 'q7_sgdea', 'q8_protecciones',
                        'q9_capacitacion', 'q10_prioridad']
        if any(not request.form.get(k, '').strip() for k in obligatorios):
            flash('Complete los campos obligatorios de identificación y las 10 preguntas principales.', 'danger')
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

        # Las respuestas condicionales y los campos adicionales se conservan en JSON.
        respuestas = {k: request.form.getlist(k) if len(request.form.getlist(k)) > 1 else request.form.get(k, '')
                      for k in request.form.keys() if k.startswith('q') or k.startswith('evidencia_')}
        respuestas.update({
            'codigo_entidad': codigo_entidad,
            'sigla_entidad': request.form.get('sigla_entidad', ''),
            'grupo_relacion': request.form.get('grupo_relacion', ''),
            'clasificacion_preliminar': request.form.get('clasificacion_preliminar', ''),
            'ambito': request.form.get('ambito', ''),
            'telefono_responsable': request.form.get('telefono_responsable', ''),
            'sitio_web': request.form.get('sitio_web', ''),
            'direccion_entidad': request.form.get('direccion_entidad', ''),
            'nit_entidad': request.form.get('nit_entidad', ''),
            'responsable_nombre': request.form.get('responsable_nombre', ''),
            'responsable_cargo': request.form.get('responsable_cargo', ''),
            'responsable_correo': correo,
            'naturaleza': request.form.get('naturaleza', ''),
            'sector': request.form.get('sector', ''),
        })
        total, aplicables, ima, nivel = calcular_diagnostico_nuevo(request.form)
        nombre_y_nit = f"{nombre} | NIT: {request.form.get('nit_entidad', '').strip()}"
        responsable = ' | '.join([request.form.get('responsable_nombre','').strip(),
                                  request.form.get('responsable_cargo','').strip(), correo])
        # Se rellenan las columnas antiguas con valores compatibles y el formulario nuevo
        # queda guardado íntegramente en respuestas_json.
        r = Censo(
            nombre_entidad=nombre_y_nit,
            codigo_entidad=codigo_entidad or None,
            naturaleza=request.form.get('naturaleza',''),
            sector=request.form.get('sector',''), responsable=responsable,
            q5=request.form.get('q1_politica','')[:200],
            q6_json=json.dumps(request.form.getlist('q6_procesos'), ensure_ascii=False),
            q7=request.form.get('q5_trd','')[:250], q8=request.form.get('q7_sgdea','')[:200],
            q9=request.form.get('q2_seguimiento','')[:250], q10=request.form.get('q9_capacitacion','')[:250],
            q11='; '.join(request.form.getlist('q6_procesos')),
            q12='; '.join(request.form.getlist('q8_protecciones')),
            q13=request.form.get('q10_prioridad',''), q14=request.form.get('q4_depositos','')[:250],
            q15=request.form.get('q2_seguimiento','')[:100], q16=request.form.get('q4_depositos','')[:200],
            q17='; '.join(request.form.getlist('q8_protecciones')), q18=request.form.get('q7_sgdea','')[:100],
            q19=request.form.get('q1_politica','')[:200], q20=request.form.get('q10_prioridad',''),
            q21=request.form.get('q9_capacitacion','')[:100],
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
            flash('Diagnóstico guardado correctamente.', 'success')
            return redirect(url_for('resultado', censo_id=r.id))
        except Exception:
            db.session.rollback()
            app.logger.exception('No fue posible guardar el diagnóstico nuevo')
            flash('No fue posible guardar el diagnóstico. Revisa los registros de la aplicación.', 'danger')
            return render_template('diagnostico.html'), 500
    return render_template('diagnostico.html')


# Equivalencias de respuestas a puntaje por componente. Se comparten entre el
# cálculo del diagnóstico y el tablero del Sistema Distrital de Archivos.
MAPAS_COMPONENTES = {
        'estrategico': [
            ('q1_politica', {'No existe': 0, 'En desarrollo': 0.5, 'Aprobada': 1,
                             'En implementación': 0.75, 'Implementada y con seguimiento': 1,
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
                        'Convalidadas': 0.75, 'En implementación': 0.75, 'Implementadas': 1,
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
    """Puntaje de 0 a 1 de cada componente (None si no hay respuestas puntuables)."""
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
    """Calcula un índice piloto con cinco componentes de igual peso (20 % c/u).

    Esta fórmula es una propuesta técnica preliminar, no una escala oficial del AGN.
    Las preguntas descriptivas no suman puntaje automáticamente.
    """
    componentes = puntajes_componentes(form)
    puntajes = [v for v in componentes.values() if v is not None]
    ima = round((sum(puntajes) / len(puntajes)) * 100, 2) if puntajes else 0
    total = round(sum(puntajes), 2)
    aplicables = float(len(puntajes))
    if ima <= 33.33: nivel = 'Bajo'
    elif ima <= 66.66: nivel = 'Medio'
    else: nivel = 'Alto'
    return total, aplicables, ima, nivel

@app.route('/resultado/<int:censo_id>')
def resultado(censo_id): return render_template('resultado.html',censo=Censo.query.get_or_404(censo_id))

@app.route('/normatividad')
def normatividad(): return render_template('normatividad.html',normatividad=NORMATIVIDAD)

@app.route('/ayuda')
def ayuda():
    return render_template('ayuda.html')



def admin_required(view):
    """Exige una sesión administrativa válida para las rutas protegidas."""
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        if not session.get("admin_authenticated"):
            return redirect(url_for("admin_login", next=request.path))
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

        if request.form.get('password', '') == configured_password:
            session.clear()
            session['admin_authenticated'] = True
            flash('Acceso administrativo autorizado.', 'success')
            destination = request.args.get('next', '')
            # Solo permitir rutas locales; impedir redirecciones externas.
            if not destination.startswith('/') or destination.startswith('//'):
                destination = url_for('admin')
            return redirect(destination)

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
    registros=Censo.query.order_by(Censo.fecha.desc()).all()
    promedio=round(sum(x.ima for x in registros)/len(registros),2) if registros else 0
    niveles={}
    sectores=sorted({(x.sector or 'Sin sector') for x in registros})
    for x in registros:
        niveles[x.nivel or 'Sin clasificar']=niveles.get(x.nivel or 'Sin clasificar',0)+1

    def semaforo(ima):
        # Clasificación simplificada en tres niveles: Bajo, Medio y Alto.
        if ima <= 33.33: return {'label':'Bajo','color':'#dc3545','class':'low'}
        if ima <= 66.66: return {'label':'Medio','color':'#e0a800','class':'medium'}
        return {'label':'Alto','color':'#198754','class':'high'}

    semaforos={
        'Bajo':sum(1 for x in registros if (x.ima or 0) <= 33.33),
        'Medio':sum(1 for x in registros if 33.33 < (x.ima or 0) <= 66.66),
        'Alto':sum(1 for x in registros if (x.ima or 0) > 66.66),
    }
    ranking=[{'nombre':x.nombre_entidad,'ima':round(x.ima or 0,2),'nivel':semaforo(x.ima or 0)['label'],'semaforo':semaforo(x.ima or 0)['label']} for x in sorted(registros,key=lambda r:r.ima or 0,reverse=True)]

    # Promedio de respuestas puntuables de cada instrumento archivístico (pregunta 6).
    instrumentos_totales={nombre:[0.0,0] for nombre in INSTRUMENTOS}
    for x in registros:
        try: respuestas=json.loads(x.q6_json or '{}')
        except (TypeError,ValueError): respuestas={}
        if not isinstance(respuestas, dict):
            respuestas = {}
        for nombre,valor in respuestas.items():
            if nombre not in instrumentos_totales: continue
            if valor in ('0','0.5','1'):
                instrumentos_totales[nombre][0]+=float(valor)
                instrumentos_totales[nombre][1]+=1
    instrumentos_labels=[]; instrumentos_valores=[]
    for nombre,(suma,cantidad) in instrumentos_totales.items():
        if cantidad:
            instrumentos_labels.append(nombre)
            instrumentos_valores.append(round(suma/cantidad*100,2))

    return render_template('admin.html',registros=registros,promedio=promedio,niveles=niveles,
        sectores=sectores,semaforos=semaforos,ranking=ranking,instrumentos_labels=instrumentos_labels,
        instrumentos_valores=instrumentos_valores,semaforo_fn=semaforo)


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
        'Puntaje obtenido', 'Variables aplicables', 'IMA (%)', 'Nivel de madurez',
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

    resumen = wb.create_sheet('Resumen IMA')
    resumen.append([
        'ID de registro', 'Fecha de registro', 'Entidad / NIT', 'Naturaleza',
        'Sector', 'Puntaje obtenido', 'Variables aplicables', 'IMA (%)', 'Nivel de madurez'
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
        'telefono_responsable', 'sitio_web', 'direccion_entidad',
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
        'Entidad / NIT', 'IMA (%)', 'Nivel', *campos_ordenados
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

# Etapas de las TRD en el orden del ciclo que acompaña el Consejo.
SDA_ETAPAS_TRD = [
    ('No elaboradas', 'Sin TRD'),
    ('En elaboración', 'En elaboración'),
    ('En evaluación', 'En evaluación'),
    ('Convalidadas', 'Convalidadas'),
    ('En implementación', 'En implementación'),
    ('Implementadas', 'Implementadas'),
]

# Temas de normalización: (clave, tema, instrumento o lineamiento sugerido, regla de brecha).
SDA_TEMAS_NORMALIZACION = [
    ('politica', 'Política de Gestión Documental', 'Guía para formular, aprobar y hacer seguimiento a la PGD',
     lambda r: r.get('q1_politica') in ('No existe', 'En desarrollo')),
    ('seguimiento', 'Seguimiento al PGD', 'Lineamiento de seguimiento y planes de mejora del PGD',
     lambda r: r.get('q2_seguimiento') == 'No'),
    ('depositos', 'Depósitos y conservación', 'Lineamiento de condiciones de depósitos y SIC',
     lambda r: r.get('q4_depositos') == 'No'),
    ('trd', 'Tablas de Retención Documental', 'Guía de elaboración y presentación de TRD',
     lambda r: r.get('q5_trd') in ('No elaboradas', 'Requieren actualización')),
    ('sgdea', 'Documento electrónico y SGDEA', 'Lineamiento de requisitos para SGDEA',
     lambda r: r.get('q7_sgdea') == 'No'),
    ('preservacion', 'Preservación digital', 'Guía del Plan de Preservación Digital a Largo Plazo',
     lambda r: 'Plan de preservación digital' not in _sda_lista(r.get('q8_protecciones'))),
    ('capacitacion', 'Capacitación y cultura archivística', 'Plan distrital de capacitación en gestión documental',
     lambda r: r.get('q9_capacitacion') == 'No se realiza'),
]

SDA_PROCESOS = ['Planeación', 'Producción', 'Gestión y trámite', 'Organización',
                'Transferencias', 'Disposición final', 'Preservación a largo plazo', 'Valoración']


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
        return 'Bajo' if ima <= 33.33 else ('Medio' if ima <= 66.66 else 'Alto')

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
    # Solo cuentan los diagnósticos del formulario actual (versiones antiguas no tienen estas preguntas).
    con_dato = [e for e in entidades if e['respuestas'].get('q1_politica')]
    temas = []
    for clave, tema, instrumento, regla in SDA_TEMAS_NORMALIZACION:
        brecha = sum(1 for e in con_dato if regla(e['respuestas']))
        temas.append({'clave': clave, 'tema': tema, 'instrumento': instrumento,
                      'entidades': brecha, 'pct': _sda_pct(brecha, len(con_dato))})
    temas.sort(key=lambda t: -t['pct'])
    procesos = []
    for proceso in SDA_PROCESOS:
        con = sum(1 for e in con_dato if proceso in _sda_lista(e['respuestas'].get('q6_procesos')))
        procesos.append({'proceso': proceso, 'entidades': con, 'pct': _sda_pct(con, len(con_dato))})
    procesos.sort(key=lambda p: p['pct'])

    # EVALUACIÓN: ciclo de las TRD.
    estados = {}
    for e in entidades:
        estado = e['respuestas'].get('q5_trd') or 'Sin dato'
        estados[estado] = estados.get(estado, 0) + 1
    etapas = [{'estado': est, 'etiqueta': et, 'entidades': estados.get(est, 0),
               'pct': _sda_pct(estados.get(est, 0), n)} for est, et in SDA_ETAPAS_TRD]
    otros_trd = {k: v for k, v in estados.items() if k not in dict(SDA_ETAPAS_TRD)}
    con_trd_avanzada = [e for e in entidades if e['respuestas'].get('q5_trd') in
                        ('Convalidadas', 'En implementación', 'Implementadas')]
    con_evidencia = sum(1 for e in con_trd_avanzada if e['respuestas'].get('q5_evidencia_estado') == 'Sí, ambas')
    requieren_actualizacion = estados.get('Requieren actualización', 0)

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
        'estados': ['Radicada', 'Asignada', 'En atención', 'Atendida', 'Cerrada'],
    },
    'evaluacion': {
        'prefijo': 'EV', 'nombre': 'Evaluación de instrumentos',
        'temas': ['TRD – presentación por primera vez', 'TRD – actualización', 'TVD'],
        'modalidades': [],
        'estados': ['Radicada', 'Revisión preliminar', 'Mesa técnica', 'Requiere ajustes',
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


def _sda_radicados_sesion():
    return list(session.get('sda_radicados', []))


@app.route('/sistema-distrital')
def sistema_distrital():
    return render_template('sistema_distrital.html')


@app.route('/servicios-en-linea')
def servicios_en_linea():
    servicio = request.args.get('servicio', '')
    return render_template('servicios_en_linea.html', tipos=SDA_TIPOS, valores={}, errores={},
                           tipo_activo=servicio if servicio in ('asistencia', 'evaluacion', 'consulta') else '')


@app.route('/sistema-distrital/solicitud', methods=['POST'])
def sistema_distrital_solicitud():
    f = request.form
    tipo = f.get('tipo', '')
    if tipo not in SDA_TIPOS:
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
    if not f.get('autoriza_datos'):
        errores['autoriza_datos'] = 'Debe autorizar el tratamiento de datos para radicar.'
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
                             descripcion=v['descripcion'], estado='Radicada')
        db.session.add(s)
        db.session.flush()
        s.radicado = f"SDA-{cfg['prefijo']}-{datetime.utcnow().year}-{s.id:05d}"
        db.session.add(SolicitudEvento(solicitud_id=s.id, estado='Radicada',
                                       observacion='Solicitud recibida por el Sistema Distrital de Archivos.'))
        db.session.commit()
    except Exception:
        db.session.rollback()
        app.logger.exception('No fue posible radicar la solicitud del Sistema Distrital')
        errores['general'] = 'No fue posible radicar la solicitud en este momento. Intente de nuevo.'
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
            error = 'No encontramos un trámite con ese radicado y correo. Verifique los datos.'
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

with app.app_context():
    db.create_all()
    # Migración compatible: agrega las nuevas columnas sin borrar registros existentes.
    # Se comprueba el esquema real porque create_all() no modifica tablas ya creadas.
    try:
        columnas = {col['name'] for col in sqlalchemy_inspect(db.engine).get_columns('censo')}
        nuevas_columnas = {
            'respuestas_json': 'TEXT',
            'codigo_entidad': 'VARCHAR(50)',
        }
        for nombre_columna, tipo_sql in nuevas_columnas.items():
            if nombre_columna not in columnas:
                with db.engine.begin() as conexion:
                    conexion.exec_driver_sql(
                        f'ALTER TABLE censo ADD COLUMN {nombre_columna} {tipo_sql}'
                    )
                app.logger.info('Columna agregada a censo: %s', nombre_columna)
    except Exception:
        app.logger.exception('No se pudo verificar/agregar una columna nueva a censo')
if __name__ == "__main__":
    app.run(debug=os.environ.get("FLASK_DEBUG", "false").lower() == "true")
