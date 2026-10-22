from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file, jsonify
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import inspect as sqlalchemy_inspect
from datetime import datetime
from pathlib import Path
from functools import wraps
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import json
import os
import csv
import unicodedata

BASE_DIR=Path(__file__).resolve().parent
app=Flask(__name__)
app.config.update(SECRET_KEY=os.environ.get("SECRET_KEY", "solo-desarrollo-cambiar-SECRET_KEY"), SQLALCHEMY_DATABASE_URI=os.environ.get("DATABASE_URL", f"sqlite:///{BASE_DIR/'diagnostico_archivistico.db'}").replace("postgres://", "postgresql://", 1), SQLALCHEMY_TRACK_MODIFICATIONS=False, SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax", SESSION_COOKIE_SECURE=os.environ.get("RENDER", "").lower() == "true")
db=SQLAlchemy(app)

# En Render configura SECRET_KEY y ADMIN_PASSWORD en Environment.
# Para producción, usa almacenamiento persistente o una base de datos administrada.

class Censo(db.Model):
    # codigo_entidad vincula cada registro al directorio oficial; respuestas_json guarda
    # las 10 preguntas nuevas y sus respuestas condicionales sin perder información.
    id=db.Column(db.Integer,primary_key=True); fecha=db.Column(db.DateTime,default=datetime.now,nullable=False)
    nombre_entidad=db.Column(db.String(300),nullable=False); naturaleza=db.Column(db.String(200),nullable=False); sector=db.Column(db.String(150),nullable=False); responsable=db.Column(db.String(300),nullable=False)
    q5=db.Column(db.String(200)); q6_json=db.Column(db.Text); q7=db.Column(db.String(250)); q8=db.Column(db.String(200)); q9=db.Column(db.String(250)); q10=db.Column(db.String(250))
    q11=db.Column(db.Text); q12=db.Column(db.Text); q13=db.Column(db.Text); q14=db.Column(db.String(250)); q15=db.Column(db.String(100)); q16=db.Column(db.String(200)); q17=db.Column(db.Text); q18=db.Column(db.String(100)); q19=db.Column(db.String(200)); q20=db.Column(db.Text); q21=db.Column(db.String(100)); respuestas_json=db.Column(db.Text); codigo_entidad=db.Column(db.String(50))
    puntaje_obtenido=db.Column(db.Float,default=0); variables_aplicables=db.Column(db.Float,default=0); ima=db.Column(db.Float,default=0); nivel=db.Column(db.String(100))

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
    if ima<=20: nivel='Nivel 1 – Inicial'
    elif ima<=40: nivel='Nivel 2 – Básico'
    elif ima<=60: nivel='Nivel 3 – En desarrollo'
    elif ima<=80: nivel='Nivel 4 – Implementado'
    else: nivel='Nivel 5 – Consolidado'
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
            db.session.commit()
            flash('Diagnóstico guardado correctamente.', 'success')
            return redirect(url_for('resultado', censo_id=r.id))
        except Exception:
            db.session.rollback()
            app.logger.exception('No fue posible guardar el diagnóstico nuevo')
            flash('No fue posible guardar el diagnóstico. Revisa los registros de la aplicación.', 'danger')
            return render_template('diagnostico.html'), 500
    return render_template('diagnostico.html')


def calcular_diagnostico_nuevo(form):
    """Calcula un índice piloto con cinco componentes de igual peso (20 % c/u).

    Esta fórmula es una propuesta técnica preliminar, no una escala oficial del AGN.
    Las preguntas descriptivas no suman puntaje automáticamente.
    """
    mapas = {
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
    componentes = {}
    for componente, preguntas in mapas.items():
        valores = []
        for campo, opciones in preguntas:
            valor = opciones.get(form.get(campo, ''))
            if valor is not None:
                valores.append(valor)
        componentes[componente] = (sum(valores) / len(valores)) if valores else None
    puntajes = [v for v in componentes.values() if v is not None]
    ima = round((sum(puntajes) / len(puntajes)) * 100, 2) if puntajes else 0
    total = round(sum(puntajes), 2)
    aplicables = float(len(puntajes))
    if ima <= 20: nivel = 'Nivel 1 – Inicial'
    elif ima <= 40: nivel = 'Nivel 2 – Básico'
    elif ima <= 60: nivel = 'Nivel 3 – En desarrollo'
    elif ima <= 80: nivel = 'Nivel 4 – Implementado'
    else: nivel = 'Nivel 5 – Consolidado'
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
    return render_template('mapa_entidades.html')


@app.route('/api/admin/mapa-entidades')
@admin_required
def api_mapa_entidades():
    """Entrega al mapa el directorio y el último diagnóstico por código de entidad.

    El código del directorio es el vínculo principal. Para registros históricos que
    no tengan código, se usa como respaldo el nombre normalizado de la entidad.
    No se inventan coordenadas: el mapa muestra una referencia municipal hasta
    que se incorporen ubicaciones verificadas por entidad.
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
    for fila in directorio:
        codigo = str(fila.get('Código', '') or '').strip()
        nombre = str(fila.get('Entidad', '') or '').strip()
        diagnostico = por_codigo.get(codigo) if codigo else None
        if diagnostico is None:
            diagnostico = por_nombre.get(clave_nombre(nombre))

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
            # Solo dibuja marcadores puntuales si el directorio incorpora
            # coordenadas explícitas y válidas; no geocodifica ni inventa lugares.
            'latitud': leer_coordenada(fila, ['Latitud', 'latitud', 'Latitude'], -90, 90),
            'longitud': leer_coordenada(fila, ['Longitud', 'longitud', 'Longitude'], -180, 180),
            'diagnostico': diagnostico
        })

    con_diagnostico = sum(1 for entidad in salida if entidad['diagnostico'] is not None)
    return jsonify({
        'entidades': salida,
        'total': len(salida),
        'con_diagnostico': con_diagnostico,
        'sin_diagnostico': len(salida) - con_diagnostico
    })


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
        # Umbrales visuales de referencia, coherentes con los cinco niveles IMA del formulario.
        if ima <= 20: return {'label':'Crítico','color':'#dc3545','class':'critical'}
        if ima <= 40: return {'label':'Bajo','color':'#fd7e14','class':'low'}
        if ima <= 60: return {'label':'En desarrollo','color':'#e0a800','class':'developing'}
        if ima <= 80: return {'label':'Satisfactorio','color':'#198754','class':'good'}
        return {'label':'Avanzado','color':'#0d6efd','class':'advanced'}

    semaforos={
        'Crítico':sum(1 for x in registros if x.ima <= 20),
        'Bajo':sum(1 for x in registros if 20 < x.ima <= 40),
        'En desarrollo':sum(1 for x in registros if 40 < x.ima <= 60),
        'Satisfactorio':sum(1 for x in registros if 60 < x.ima <= 80),
        'Avanzado':sum(1 for x in registros if x.ima > 80),
    }
    ranking=[{'nombre':x.nombre_entidad,'ima':round(x.ima or 0,2),'nivel':x.nivel or 'Sin clasificar','semaforo':semaforo(x.ima or 0)['label']} for x in sorted(registros,key=lambda r:r.ima or 0,reverse=True)]

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
