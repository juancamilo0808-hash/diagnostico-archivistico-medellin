"""Cuestionario del Diagnóstico Territorial Archivístico, versión V04.

Fuente: documento «Preguntas diagnóstico V04» del Consejo Distrital de Archivos.
El formulario, el puntaje, la ficha, los riesgos y la exportación se generan a partir
de estos datos: para cambiar una pregunta, una opción o una pregunta hija basta con
editar este archivo.

Puntaje: cada pregunta evaluativa vale No = 0, Parcialmente = 0,5 y Sí = 1. «N/A – No
aplica» no se cuenta. El puntaje de cada componente es el promedio de sus preguntas y
el nivel de madurez es el promedio de los cinco componentes (20 % cada uno).
Las preguntas de caracterización (C1–C9) y las preguntas hijas son informativas.
"""

VERSION = 'V04'

PUNTOS = {'no': 0.0, 'parcial': 0.5, 'si': 1.0, 'na': None}

# ----- Opciones de respuesta -----------------------------------------------------
NO = ('no', 'No')
SI = ('si', 'Sí')
PARCIAL = ('parcial', 'Parcialmente')
NA = ('na', 'N/A – No aplica')
EN_PROCESO = ('parcial', 'En proceso de elaboración')

SN = [NO, SI]
SNP = [NO, SI, PARCIAL]
SNP_NA = [NO, SI, PARCIAL, NA]


# ----- Preguntas hijas ---------------------------------------------------------------
def url(etiqueta, cuando=('si', 'parcial')):
    """Enlace de evidencia: campo URL con la opción «Enlace no disponible»."""
    return {'id': 'url', 'tipo': 'url', 'etiqueta': etiqueta, 'cuando': cuando}


def justificacion(cuando=('parcial',), etiqueta='Justificación', requerido=True):
    return {'id': 'just', 'tipo': 'textarea', 'etiqueta': etiqueta, 'cuando': cuando,
            'requerido': requerido, 'max': 500}


def just_na(etiqueta='Justificación de N/A'):
    return {'id': 'just_na', 'tipo': 'textarea', 'etiqueta': etiqueta, 'cuando': ('na',),
            'requerido': True, 'max': 500}


def sel(id_, etiqueta, opciones, cuando=('si', 'parcial')):
    return {'id': id_, 'tipo': 'select', 'etiqueta': etiqueta, 'opciones': opciones, 'cuando': cuando}


def multi(id_, etiqueta, opciones, cuando=('si', 'parcial'), otro=None):
    """Selección múltiple. `otro` es la opción que habilita el campo «¿Cuál?»."""
    return {'id': id_, 'tipo': 'multi', 'etiqueta': etiqueta, 'opciones': opciones,
            'cuando': cuando, 'otro': otro}


def texto(id_, etiqueta, cuando=('si', 'parcial'), max_=500):
    return {'id': id_, 'tipo': 'textarea', 'etiqueta': etiqueta, 'cuando': cuando, 'max': max_}


def pregunta(n, texto_, opciones, criterio, tema, hijas=(), ayuda=None, excluir=None):
    return {'n': n, 'campo': f'p{n}', 'texto': texto_, 'opciones': list(opciones),
            'criterio': criterio, 'tema': tema, 'hijas': list(hijas), 'ayuda': ayuda,
            'excluir': excluir}


PROGRAMAS_PGD = [
    'Programa de normalización de formas y formularios electrónicos',
    'Programa de documentos vitales o esenciales',
    'Programa de gestión de documentos electrónicos',
    'Programa de archivos descentralizados',
    'Programa de reprografía',
    'Programa de documentos especiales',
    'Plan Institucional de Capacitación',
    'Programa de auditoría y control',
]
ESTADOS_PROGRAMA = ['Formulado e implementado', 'Formulado sin implementación', 'En formulación',
                    'No identificado como necesario']
ESTADOS_PARCIAL_POLITICA = ['Elaborada y no aprobada', 'Elaborada, aprobada y desactualizada',
                            'En proceso de elaboración']
INSTRUMENTOS_ARCHIVISTICOS = [
    'Registro de Activos de Información (RAI)',
    'Índice de Información Clasificada y Reservada (IICR)',
    'Esquema de Publicación de Información (EPI)',
    'Programa de Gestión Documental (PGD)',
    'Tablas de Retención Documental (TRD)',
    'Cuadro de Clasificación Documental (CCD)',
    'Tablas de Valoración Documental (TVD)',
    'Plan Institucional de Archivos (PINAR)',
    'Formato Único de Inventario Documental (FUID)',
    'Tablas de Control de Acceso (TCA)',
    'Bancos Terminológicos (BANTER)',
    'Modelo de Requisitos para la Gestión de Documentos Electrónicos (MoReq)',
    'Sistema Integrado de Conservación (SIC)',
    'Mapas de procesos, flujos documentales y descripción de funciones de las unidades administrativas',
]
ESTADOS_INSTRUMENTO_TRD = ['No cuenta', 'Elaborada', 'Aprobada', 'Convalidada', 'Implementada']


# ----- Caracterización (C1–C9): informativa, sin puntaje -----------------------------
NATURALEZAS = ['Pública', 'Privada con función pública', 'Mixta', 'Privada']
SECTORES = ['Sector central', 'Sector descentralizado',
            'Entidades indirectas (sociedades, asociaciones, cooperativas y fundaciones públicas y mixtas)',
            'Otro']
NOTA_CARACTERIZACION = (
    'Estas preguntas recopilan información general sobre la naturaleza jurídica, la estructura '
    'administrativa, los datos de contacto, los responsables de la gestión documental y las '
    'características de los archivos de la entidad. Son informativas: no generan puntuación ni '
    'inciden en el nivel de madurez archivística.')


# ----- Componentes evaluativos -------------------------------------------------------
COMPONENTES = [
    {'clave': 'estrategico', 'num': 1, 'nombre': 'Estratégico', 'peso': 20, 'subcomponentes': [
        {'nombre': 'Planeación de la función archivística',
         'preguntas': [
            pregunta(1, '¿La entidad tiene formalizadas las responsabilidades de gestión documental?', SN,
                     'Acto, manual, mapa de procesos o instrumento equivalente vigente.',
                     'Responsabilidades de gestión documental formalizadas'),
            pregunta(2, '¿La entidad cuenta con una política de gestión documental aprobada y vigente?', SNP,
                     'Política formalizada y actualizada; la publicación y alcance se verifican como evidencia.',
                     'Política de gestión documental aprobada y vigente',
                     hijas=[sel('estado', 'Estado de la política', ESTADOS_PARCIAL_POLITICA, ('parcial',)),
                            url('Enlace (URL) donde se puede consultar la política institucional de archivos y gestión documental y su acto de aprobación')]),
            pregunta(3, '¿La entidad dispone de un Diagnóstico Integral de Archivos (DIA) actualizado?', [NO, SI, EN_PROCESO],
                     'Documento del Diagnóstico Integral de Archivos.',
                     'Diagnóstico Integral de Archivos actualizado',
                     hijas=[url('Enlace (URL) donde se puede consultar el diagnóstico integral de archivos vigente')]),
            pregunta(4, '¿La entidad dispone de los instrumentos archivísticos, según el Decreto 1080 de 2015 y el Acuerdo 001 de 2024?', [NO, SI, EN_PROCESO],
                     'Instrumentos archivísticos y sus actos de aprobación.',
                     'Instrumentos archivísticos (Decreto 1080 de 2015 y Acuerdo 001 de 2024)',
                     ayuda=INSTRUMENTOS_ARCHIVISTICOS,
                     hijas=[url('Enlace (URL) donde se pueden consultar los instrumentos y el acto administrativo de aprobación')]),
         ]},
        {'nombre': 'Planeación estratégica',
         'preguntas': [
            pregunta(5, '¿El Programa de Gestión Documental (PGD) se encuentra elaborado, aprobado, vigente y en ejecución?', SNP,
                     'PGD aprobado, vigente y en ejecución.',
                     'Programa de Gestión Documental (PGD) aprobado y en ejecución',
                     hijas=[sel('estado', 'Estado del PGD', ESTADOS_PARCIAL_POLITICA + ['Aprobado e implementado'], ('parcial',)),
                            url('Enlace (URL) donde se puede consultar el Programa de Gestión Documental (PGD) aprobado y vigente')]),
            pregunta(6, '¿El Plan Institucional de Archivos (PINAR) se encuentra actualizado en la última vigencia?', SN,
                     'PINAR.', 'Plan Institucional de Archivos (PINAR) actualizado',
                     hijas=[url('Enlace (URL) donde se puede consultar el Plan Institucional de Archivos (PINAR)', ('si',))]),
            pregunta(7, '¿La entidad articula la Política de Gestión Documental con la planeación institucional y los sistemas de gestión de la entidad?', SNP,
                     'Instrumentos de planeación y políticas con los que se articula la gestión documental.',
                     'Articulación de la política con la planeación y los sistemas de gestión',
                     hijas=[multi('instrumentos', '¿Con cuáles instrumentos de planeación o políticas aplicables está articulada la gestión documental?',
                                  ['Planeación institucional', 'Gestión de riesgos', 'Transparencia', 'Seguridad digital',
                                   'Gestión del conocimiento', 'Otra'], otro='Otra')]),
            pregunta(8, '¿El PGD de la entidad identifica y desarrolla los programas específicos requeridos según sus necesidades y características documentales?', SNP,
                     'Programas específicos documentados.', 'Programas específicos del PGD',
                     hijas=[{'id': 'prog', 'tipo': 'matriz', 'etiqueta': 'Estado de cada programa específico',
                             'filas': PROGRAMAS_PGD, 'opciones': ESTADOS_PROGRAMA, 'cuando': ('si', 'parcial')},
                            texto('prog_just', 'Justificación de los programas no identificados como necesarios'),
                            url('Enlace al PGD y a la sección de programas específicos')]),
         ]},
        {'nombre': 'Control, evaluación y seguimiento',
         'preguntas': [
            pregunta(9, '¿La entidad evalúa los resultados de la Política de Gestión Documental y utiliza sus hallazgos para formular y establecer acciones de mejora institucional?', SN,
                     'Indicadores, actas, responsables, plazos y acciones.',
                     'Evaluación de la política y acciones de mejora'),
            pregunta(10, '¿Los riesgos de gestión documental están identificados y cuentan con medidas de tratamiento y seguimiento?', SNP,
                     'Matriz de riesgos y seguimiento.', 'Riesgos de gestión documental identificados y tratados',
                     hijas=[multi('medidas', '¿Cuáles son las medidas de tratamiento y seguimiento?',
                                  ['Plan de mejora', 'Matriz de riesgos', 'Matriz de seguimiento', 'Otra'], otro='Otra')]),
            pregunta(11, '¿La entidad ejecuta actividades de auditoría y control sobre la gestión documental y realiza seguimiento a los hallazgos y acciones de mejora?', SNP,
                     'Programa específico o mecanismo equivalente, informes, hallazgos y cierre de acciones.',
                     'Auditoría y control sobre la gestión documental',
                     hijas=[sel('tipo_si', 'Mecanismo', ['Auditorías y seguimiento realizados', 'Plan de mejoramiento'], ('si',)),
                            sel('tipo_parcial', 'Situación', ['Auditorías realizadas sin seguimiento completo',
                                                              'Programa formulado sin ejecución', 'En formulación'], ('parcial',)),
                            url('Enlace al programa, informes de auditoría o plan de mejoramiento')]),
         ]},
    ]},
    {'clave': 'administracion_archivos', 'num': 2, 'nombre': 'Administración de archivos', 'peso': 20, 'subcomponentes': [
        {'nombre': 'Administración',
         'preguntas': [
            pregunta(12, '¿La entidad cuenta con Comité Institucional de Gestión y Desempeño y/o Comité de Archivo?', SN,
                     'Acto administrativo del comité o comités.', 'Comité Institucional de Gestión y Desempeño o Comité de Archivo'),
            pregunta(13, '¿La entidad asigna recursos financieros para la ejecución y desarrollo de las actividades de gestión documental?',
                     [NO, SI, ('parcial', 'Parcialmente, porque no cubre todas las necesidades')],
                     'Presupuesto institucional, planes de inversión, contratos, otros.',
                     'Recursos financieros para la gestión documental',
                     hijas=[multi('fuentes', 'Fuente de los recursos',
                                  ['Presupuesto institucional', 'Planes de inversión', 'Contratos', 'Otros'], otro='Otros')]),
            pregunta(14, '¿La entidad cuenta con servicios tercerizados de custodia y conservación de los documentos?', SNP,
                     'Informe de verificación de seguimiento.', 'Servicios tercerizados de custodia y conservación',
                     hijas=[sel('seguimiento', '¿Realiza seguimiento técnico a los servicios de custodia y conservación que tiene tercerizados?', ['Sí', 'No'])]),
            pregunta(15, '¿La entidad ha definido y aplica controles archivísticos para los archivos ubicados en diferentes sedes, dependencias o servicios de custodia, cuando esta modalidad existe?', SNP,
                     'Inventario de ubicaciones, responsabilidades, condiciones de custodia y seguimiento.',
                     'Controles para archivos en distintas sedes o servicios de custodia',
                     excluir=('descentralizados', 'No'),
                     hijas=[sel('descentralizados', '¿Existen archivos descentralizados?', ['Sí', 'No'], ('no', 'si', 'parcial')),
                            sel('controles', 'Controles', ['Controles aplicados en todas las ubicaciones', 'Controles aplicados parcialmente',
                                                          'Controles documentados sin ejecución', 'Sin controles']),
                            url('Enlace al programa de archivos descentralizados, procedimiento de administración o registros de supervisión')]),
         ]},
        {'nombre': 'Recursos físicos',
         'preguntas': [
            pregunta(16, '¿Los espacios de archivo cuentan con condiciones adecuadas de almacenamiento, estantería y control de acceso?', SNP,
                     'Evidencias de verificación o seguimiento de las condiciones físicas.',
                     'Condiciones de almacenamiento, estantería y control de acceso'),
            pregunta(17, '¿La entidad aplica medidas de conservación preventiva y de prevención y atención de emergencias para sus documentos físicos?', SNP,
                     'Sistema Integrado de Conservación (SIC), plan y acciones de conservación y emergencias según riesgos.',
                     'Conservación preventiva y atención de emergencias'),
            pregunta(18, '¿La entidad cuenta con depósitos de archivo con la capacidad de almacenamiento suficiente de acuerdo a sus necesidades?', SNP,
                     'Informe de depósitos de archivo y su capacidad.', 'Capacidad de los depósitos de archivo',
                     hijas=[multi('capacidad', 'Capacidad de recepción',
                                  ['Con capacidad de recepción', 'Sin capacidad de recepción', 'Otros'], otro='Otros')]),
         ]},
        {'nombre': 'Talento humano',
         'preguntas': [
            pregunta(19, '¿El personal encargado de gestión documental cuenta con perfiles adecuados y recibe capacitación periódicamente?', SNP,
                     'Perfiles, plan o registros de capacitación.', 'Perfiles y capacitación del personal de gestión documental',
                     hijas=[justificacion()]),
            pregunta(20, '¿La entidad dispone de mecanismos para asegurar la entrega de archivos y la continuidad del conocimiento archivístico cuando cambia el personal?', SNP_NA,
                     'Inventarios de entrega y registros de copias de seguridad, acta de entrega de puesto de trabajo o empalme.',
                     'Entrega de archivos y continuidad del conocimiento',
                     hijas=[just_na()]),
            pregunta(21, '¿Se realizan capacitaciones al personal administrativo y misional en gestión documental articulada con la planeación institucional del talento humano?', SNP,
                     'Plan de capacitación, registros de participación y evaluación.',
                     'Capacitación del personal en gestión documental', hijas=[justificacion()]),
         ]},
        {'nombre': 'Gestión en seguridad y salud ocupacional',
         'preguntas': [
            pregunta(22, '¿Las actividades y espacios de archivo se incluyen en la identificación y control de riesgos de seguridad y salud en el trabajo?', SNP,
                     'Evidencia del SG-SST aplicable a tareas archivísticas; revisar alcance específico.',
                     'Riesgos de seguridad y salud en el trabajo en el archivo', hijas=[justificacion()]),
            pregunta(23, '¿La entidad suministra al personal que desarrolla actividades de gestión documental y archivo los elementos de protección personal (EPP) requeridos, de acuerdo con los riesgos identificados, y verifica su uso adecuado?', SNP,
                     'Matriz de peligros y valoración de riesgos, registros de entrega de EPP, capacitación en su utilización y evidencias de seguimiento.',
                     'Elementos de protección personal (EPP)',
                     hijas=[multi('epp', '¿Qué elementos de protección personal suministra la entidad?',
                                  ['Guantes de protección', 'Protección respiratoria según el riesgo', 'Gafas de protección',
                                   'Ropa de protección', 'Calzado de seguridad', 'Otros'], otro='Otros')]),
            pregunta(24, '¿La entidad garantiza que el personal que realiza actividades archivísticas consideradas trabajo en alturas cuenta con la capacitación, autorización y medidas de prevención y protección exigidas por la normativa vigente?', SNP_NA,
                     'Identificación de tareas y riesgos, certificados de formación cuando sean exigibles, procedimientos de trabajo seguro, registros de inspección y medidas de prevención y protección, curso de alturas cuando aplique.',
                     'Trabajo en alturas en actividades archivísticas',
                     hijas=[multi('medidas', '¿Qué medidas implementa la entidad para desarrollar actividades archivísticas que constituyan trabajo en alturas?',
                                  ['Capacitación y entrenamiento exigibles', 'Autorización del personal', 'Equipos de acceso y protección adecuados',
                                   'Procedimientos documentados', 'Inspección y mantenimiento de equipos', 'Otras'], otro='Otras'),
                            just_na('Justificación de N/A (la entidad no desarrolla actividades archivísticas que constituyan trabajo en alturas)')]),
         ]},
    ]},
    {'clave': 'procesos_documentales', 'num': 3, 'nombre': 'Procesos de la gestión documental', 'peso': 20, 'subcomponentes': [
        {'nombre': 'Planeación técnica',
         'preguntas': [
            pregunta(25, '¿La entidad cuenta con Tablas de Retención Documental (TRD) y/o Tablas de Valoración Documental (TVD), según corresponda, actualizadas y aprobadas por la instancia competente?', SNP_NA,
                     'Instrumentos y actos de aprobación; verificar convalidación cuando aplique.',
                     'Tablas de Retención y de Valoración Documental actualizadas y aprobadas',
                     hijas=[sel('trd', 'Estado de la TRD', ESTADOS_INSTRUMENTO_TRD),
                            sel('tvd', 'Estado de la TVD', ESTADOS_INSTRUMENTO_TRD),
                            url('Enlace (URL) donde se pueden consultar las TRD o TVD y sus actos de aprobación y convalidación, cuando corresponda'),
                            just_na()]),
            pregunta(26, '¿La entidad cuenta con Tablas de Control de Acceso u otro instrumento documentado que establezca permisos y restricciones de acceso a documentos y expedientes según roles y disposiciones legales?', SNP,
                     'Tablas de control de acceso.', 'Tablas de Control de Acceso',
                     hijas=[sel('estado', 'Estado', ['Formalizado y aplicado', 'Formalizado con aplicación parcial', 'En elaboración']),
                            url('Enlace a las TCA o matriz de permisos, sin exponer información protegida')]),
            pregunta(27, '¿La entidad dispone de un banco terminológico de series, subseries y tipos documentales y lo utiliza para normalizar la denominación y descripción de sus documentos?',
                     [NO, ('si', 'Sí, elaborado, actualizado y aplicado')], 'BANTER.', 'Banco terminológico (BANTER)'),
            pregunta(28, '¿La entidad define requisitos archivísticos, administrativos, legales y tecnológicos antes de crear o modificar documentos, formatos, expedientes y procedimientos documentales?', SN,
                     'Proceso y/o procedimiento de planeación documental o lineamientos de diseño.',
                     'Requisitos para crear o modificar documentos y procedimientos',
                     hijas=[url('Enlace al procedimiento de planeación documental o lineamientos de diseño', ('si',))]),
         ]},
        {'nombre': 'Producción',
         'preguntas': [
            pregunta(29, '¿La producción de documentos oficiales se realiza conforme a criterios institucionales de formatos, identificación y control documental?', SN,
                     'Procedimiento de producción documental, manual de comunicaciones oficiales, formatos institucionales normalizados y muestras de documentos oficiales internos y externos.',
                     'Producción de documentos oficiales con criterios institucionales',
                     ayuda='Tales como memorandos, actas, informes, circulares, resoluciones y demás documentos generados en desarrollo de las funciones institucionales; además oficios, comunicaciones oficiales, respuestas a peticiones, certificaciones, notificaciones y documentos dirigidos a ciudadanos, usuarios u otras entidades.'),
            pregunta(30, '¿La entidad aplica criterios de autenticidad, integridad, identificación, fiabilidad, disponibilidad y control de versiones en la creación de documentos electrónicos?', SNP,
                     'Programa específico de normalización de formas y formularios electrónicos.',
                     'Criterios para la creación de documentos electrónicos', hijas=[justificacion()]),
            pregunta(31, '¿Los formatos utilizados en la entidad se encuentran aprobados por el Sistema de Gestión de Calidad?', SN,
                     'Listado de formatos codificados en el Sistema de Gestión de la Calidad.',
                     'Formatos aprobados por el Sistema de Gestión de Calidad'),
         ]},
        {'nombre': 'Gestión y trámite',
         'preguntas': [
            pregunta(32, '¿La entidad controla la recepción, radicación, distribución y seguimiento de sus comunicaciones oficiales?', SNP,
                     'Registros de radicación y control de comunicaciones oficiales.',
                     'Control de comunicaciones oficiales', hijas=[justificacion()]),
            pregunta(33, '¿La entidad dispone de mecanismos documentados para consultar, prestar y recuperar expedientes, y controlar su devolución o cierre, preservando las restricciones legales de acceso?', SN,
                     'Registros de consulta y préstamo.', 'Consulta y préstamo de expedientes'),
         ]},
        {'nombre': 'Organización',
         'preguntas': [
            pregunta(34, '¿Los expedientes se conforman y organizan conforme a las series documentales y al principio de procedencia?', SN,
                     'Imágenes del almacenamiento y conservación de los expedientes.',
                     'Expedientes organizados por series y principio de procedencia'),
            pregunta(35, '¿La entidad aplica inventarios documentales y controles de foliación u ordenación, según el soporte, para localizar y verificar expedientes?', SNP,
                     'URL del FUID.', 'Inventarios documentales y foliación',
                     hijas=[sel('situacion', 'Situación', ['Inventariado y no foliado', 'Foliado y no inventariado'], ('parcial',)),
                            url('Enlace (URL) del Formato Único de Inventario Documental (FUID)')]),
            pregunta(36, '¿La entidad aplica criterios normalizados de descripción documental que permitan identificar, localizar y recuperar sus unidades documentales, más allá de la elaboración de inventarios?', SN,
                     'Hoja de control, índices electrónicos.', 'Descripción documental normalizada'),
         ]},
        {'nombre': 'Transferencia',
         'preguntas': [
            pregunta(37, '¿La entidad realiza transferencias documentales primarias y/o secundarias conforme a un procedimiento y con inventarios verificados?', SNP,
                     'Cronogramas, plan de transferencias, actas e inventarios de transferencias aplicables.',
                     'Transferencias documentales',
                     hijas=[justificacion(),
                            multi('tipo', '¿Qué tipo de transferencia documental se ha ejecutado en la última vigencia?',
                                  ['Primaria', 'Secundaria', 'De documentos electrónicos', 'Ninguna durante la vigencia']),
                            url('Enlace (URL) del procedimiento y del acta o inventario de una transferencia documental realizada')]),
            pregunta(38, '¿La entidad verifica y documenta la integridad, correspondencia y condiciones de los expedientes recibidos en las transferencias documentales?', SN,
                     'Imágenes de evidencias.', 'Verificación de expedientes recibidos en transferencias',
                     hijas=[sel('verifica', 'Verificación', ['Verifica', 'No verifica'], ('si',))]),
         ]},
        {'nombre': 'Disposición de documentos',
         'preguntas': [
            pregunta(39, '¿Las decisiones de disposición final se ejecutan conforme a las TRD o TVD aplicables y con soportes documentados?', SN,
                     'Inventarios de los documentos a eliminar, comprobante de publicación de inventarios a eliminar en la página web, comprobante de picado certificado cumpliendo con lineamientos del Sistema de Gestión Ambiental (si aplica), actas y procedimientos de eliminación.',
                     'Disposición final conforme a TRD o TVD'),
         ]},
        {'nombre': 'Preservación a largo plazo',
         'preguntas': [
            pregunta(40, '¿La entidad dispone de un Sistema Integrado de Conservación (SIC) y se implementa?', SNP,
                     'SIC, planes de conservación y preservación, registros de ejecución.',
                     'Sistema Integrado de Conservación (SIC)',
                     hijas=[sel('plan', '¿Cuál de los dos planes tiene e implementa?',
                                ['Plan de Conservación Documental', 'Plan de Preservación Digital a Largo Plazo'], ('parcial',))]),
            pregunta(41, '¿La entidad ha identificado documentos en soportes o formatos especiales y aplica medidas diferenciadas para su organización, conservación, acceso y preservación?', SN,
                     'Programa específico de documentos especiales, inventarios, procedimientos por soporte y registros de tratamiento.',
                     'Documentos en soportes o formatos especiales',
                     hijas=[sel('estado', 'Estado', ['Identificados y gestionados con medidas diferenciadas',
                                                     'Identificados con medidas parciales', 'En identificación'], ('si',)),
                            multi('soportes', '¿Cuáles?', ['Fotográfico', 'Audiovisual', 'Sonoro', 'Cartográfico', 'Otros'],
                                  ('si',), otro='Otros')]),
            pregunta(42, '¿La entidad ha definido y aplica criterios técnicos para la reproducción, digitalización o microfilmación de documentos, asegurando calidad, identificación y control de las copias?', SNP,
                     'Programa específico de reprografía, evidencias de implementación.',
                     'Criterios de reprografía y digitalización',
                     hijas=[sel('estado', 'Estado', ['Criterios definidos y aplicados', 'Aplicación parcial', 'En formulación']),
                            multi('tecnologias', 'Tecnologías utilizadas', ['Digitalización', 'Microfilmación', 'Fotocopiado', 'Otras'], otro='Otras')]),
            pregunta(43, '¿La entidad identifica sus documentos vitales o esenciales y aplica medidas específicas para protegerlos y recuperarlos ante emergencias o interrupciones?', SN,
                     'Programa específico de documentos vitales o esenciales, informes de implementación.',
                     'Documentos vitales o esenciales'),
         ]},
    ]},
    {'clave': 'tecnologico', 'num': 4, 'nombre': 'Tecnológico', 'peso': 20, 'subcomponentes': [
        {'nombre': 'Articulación de la gestión de documentos electrónicos',
         'preguntas': [
            pregunta(44, '¿El Sistema de Gestión de Documentos Electrónicos de Archivo (SGDEA) de la entidad se encuentra articulado con la Política de Archivos y Gestión Documental, el PGD, el PINAR, el SIC, las TRD, las TVD y los demás instrumentos archivísticos aplicables, según las prioridades institucionales y las responsabilidades de Gestión Documental y Tecnologías de la Información?', SNP,
                     'Manual de procedimiento del SGDEA.', 'SGDEA articulado con los instrumentos archivísticos',
                     hijas=[multi('instrumentos', 'Instrumentos articulados con el SGDEA',
                                  ['Cuadro de Clasificación Documental (CCD)', 'Tablas de Retención Documental (TRD)',
                                   'Programa de Gestión Documental (PGD)', 'Plan Institucional de Archivos (PINAR)',
                                   'Formato Único de Inventario Documental (FUID)', 'Tablas de Control de Acceso (TCA)',
                                   'Banco Terminológico (BANTER)', 'Índices de Gestión de la Información Pública (IGIP)',
                                   'Sistema Integrado de Conservación (SIC)', 'Tablas de Valoración Documental (TVD)',
                                   'Modelo de Requisitos (MoReq)'])]),
         ]},
        {'nombre': 'Tecnologías para la gestión de documentos electrónicos',
         'preguntas': [
            pregunta(45, '¿El Sistema de Gestión de Documentos Electrónicos de Archivo (SGDEA) de la entidad cuenta con todos los parámetros establecidos por el Modelo de Requisitos (MoReq)?', SNP,
                     'MoReq, manual de procesos y procedimientos del SGDEA.', 'SGDEA conforme al Modelo de Requisitos (MoReq)',
                     hijas=[justificacion()]),
            pregunta(46, '¿Los documentos electrónicos de archivo se capturan e incorporan a expedientes electrónicos según la TRD, con índice electrónico, disposición final, metadatos definidos en un esquema aprobado y mecanismos de autenticidad e integridad?', SNP,
                     'Esquema de metadatos aprobado, expediente con índice electrónico y mecanismos de autenticidad.',
                     'Expedientes electrónicos con índice, metadatos y autenticidad',
                     hijas=[multi('elementos', '¿Qué elementos aplica la entidad?',
                                  ['Índice electrónico', 'Disposición final según la TRD', 'Metadatos en un esquema aprobado',
                                   'Mecanismos de autenticidad e integridad'])]),
            pregunta(47, '¿La entidad ejecuta las actividades del Programa Específico de Gestión de Documentos Electrónicos previstas en su Programa de Gestión Documental (PGD)?', SNP,
                     'Avance de actividades del programa, informes de seguimiento.',
                     'Programa específico de gestión de documentos electrónicos',
                     hijas=[sel('estado', 'Estado', ['Programa en ejecución y seguimiento', 'Ejecución parcial',
                                                     'Formulado sin ejecución', 'En formulación'], ('parcial',))]),
         ]},
        {'nombre': 'Seguridad y privacidad',
         'preguntas': [
            pregunta(48, '¿Los sistemas que gestionan documentos físicos y electrónicos de archivo aplican controles de acceso según roles y la clasificación de la información (pública, clasificada o reservada)?', SN,
                     'Índice de información clasificada y reservada, tablas de control de acceso.',
                     'Controles de acceso según roles y clasificación de la información'),
            pregunta(49, '¿La entidad realiza respaldos de los documentos electrónicos de archivo y cuenta con procedimientos probados de recuperación, respuesta a incidentes y continuidad ante pérdida, alteración, acceso no autorizado, indisponibilidad o código malicioso?', SN,
                     'Procedimiento de copias de seguridad y respaldo.', 'Respaldos y recuperación de documentos electrónicos',
                     hijas=[sel('frecuencia', '¿Con qué frecuencia se hacen los respaldos?', ['Diaria', 'Semanal', 'Mensual', 'Otra'], ('si',)),
                            texto('donde', '¿Dónde se almacena la información del respaldo?', ('si',))]),
            pregunta(50, '¿La entidad verifica periódicamente la integridad y disponibilidad de sus documentos y expedientes electrónicos, y controla sus formatos y el riesgo de obsolescencia conforme al Plan de Preservación Digital a Largo Plazo?', SNP,
                     'Informes de seguimiento e implementación del plan de preservación.',
                     'Integridad, formatos y obsolescencia (preservación digital)',
                     hijas=[multi('acciones', 'Acciones que realiza',
                                  ['Verificación de integridad', 'Identificación y control de formatos', 'Migración o normalización',
                                   'Copias en ubicaciones distintas', 'Repositorio de preservación', 'Otra'], otro='Otra')]),
            pregunta(51, '¿La entidad realiza seguimiento y evaluación periódica de la gestión de documentos electrónicos y del SGDEA, con indicadores o informes que permitan corregir desviaciones y planear mejoras?', SN,
                     'Informes, indicadores y planes de mejora.', 'Seguimiento a la gestión de documentos electrónicos y al SGDEA'),
         ]},
        {'nombre': 'Interoperabilidad',
         'preguntas': [
            pregunta(52, '¿La entidad identifica los sistemas de información que producen o gestionan documentos de archivo y, cuando existe intercambio entre ellos (trámites compartidos), dispone de mecanismos de interoperabilidad para transferir documentos y metadatos conservando identificación, estructura, contexto, integridad y trazabilidad?', SNP,
                     'Inventario de sistemas y relación de cuáles son interoperables entre sí.',
                     'Interoperabilidad de los sistemas de información', hijas=[justificacion()]),
            pregunta(53, '¿La entidad documenta y valida las migraciones o transferencias de documentos entre plataformas para evitar pérdida de estructura, obsolescencia, metadatos o contexto?', SNP,
                     'Protocolos de migración, controles de integridad y actas de validación.',
                     'Migraciones documentadas y validadas', hijas=[justificacion()]),
         ]},
    ]},
    {'clave': 'cultural', 'num': 5, 'nombre': 'Cultural', 'peso': 20, 'subcomponentes': [
        {'nombre': 'Gestión del conocimiento',
         'preguntas': [
            pregunta(54, '¿La entidad utiliza los documentos y archivos para la gestión del conocimiento, la innovación y la memoria pública, articulándolos con las políticas institucionales aplicables?', SN,
                     'Informe de seguimiento e implementación a las políticas institucionales.',
                     'Archivos para la gestión del conocimiento y la memoria pública'),
            pregunta(55, '¿La entidad desarrolla actividades de sensibilización o formación sobre la importancia de los archivos y la memoria institucional?', SN,
                     'Asistencias a las capacitaciones.', 'Sensibilización sobre archivos y memoria institucional'),
            pregunta(56, '¿La entidad evalúa los resultados de la capacitación específica en gestión documental y ajusta sus contenidos según las necesidades detectadas?', SN,
                     'Evaluaciones de capacitaciones.', 'Evaluación de la capacitación en gestión documental'),
         ]},
        {'nombre': 'Patrimonio documental',
         'preguntas': [
            pregunta(57, '¿La entidad ha identificado e inventariado los fondos o documentos de especial valor histórico, cultural, social, científico, de Derechos Humanos (DDHH), Derecho Internacional Humanitario (DIH) o de conservación permanente bajo su responsabilidad?', SNP,
                     'Inventarios, estudios de valoración patrimonial, TRD, TVD, CCD y soportes de identificación.',
                     'Fondos de valor patrimonial, DDHH o DIH identificados e inventariados',
                     hijas=[multi('situacion', 'Situación',
                                  ['Cuentan con estudio de valoración patrimonial formalizado', 'Están identificados en TRD, TVD o CCD',
                                   'Identificados empíricamente sin acto administrativo de valoración', 'En proceso de identificación'],
                                  ('parcial',))]),
            pregunta(58, '¿La entidad tiene criterios documentados para identificar, valorar y gestionar fondos documentales con posible interés patrimonial?', SNP,
                     'Criterios de identificación, inventarios y registros de valoración.',
                     'Criterios para valorar fondos con interés patrimonial', hijas=[justificacion()]),
            pregunta(59, '¿La entidad aplica medidas de protección para documentos de valor patrimonial o relacionados con derechos humanos, memoria y paz, cuando corresponde?', SNP,
                     'Medidas de protección, acceso y conservación según naturaleza documental.',
                     'Protección de documentos patrimoniales y de derechos humanos', hijas=[justificacion()]),
            pregunta(60, '¿La entidad ha adelantado gestiones o trámites de postulación, inclusión o declaratoria de sus acervos patrimoniales en la Lista Indicativa de Candidatos a Bienes de Interés Cultural (LICBIC-CDA) o en el Registro BIC-CDA ante las autoridades competentes?', SNP,
                     'Resoluciones, actos de declaratoria, constancias de inclusión o radicados de postulación.',
                     'Trámites de LICBIC-CDA o declaratoria BIC-CDA',
                     hijas=[sel('estado', 'Estado del trámite', ['Cuenta con bienes declarados BIC-CDA', 'Incluidos en LICBIC-CDA',
                                                               'Trámite iniciado o en curso', 'No ha iniciado trámites']),
                            url('Enlace (URL) donde se puede consultar el acto, registro, constancia o radicado del trámite BIC-CDA')]),
            pregunta(61, '¿La entidad ha formulado, aprobado o implementado un Plan Especial de Manejo y Protección de Carácter Documental Archivístico (PEMP-CDA), cuando corresponde?', SNP,
                     'Documento del plan, acto de aprobación, informes de implementación o soportes de formulación.',
                     'Plan Especial de Manejo y Protección (PEMP-CDA)',
                     hijas=[sel('estado', 'Estado del PEMP-CDA', ['Aprobado e implementado', 'Aprobado sin implementación completa',
                                                                 'En formulación o actualización', 'No cuenta con PEMP-CDA'])]),
            pregunta(62, '¿La entidad dispone de personal con competencias especializadas para gestionar fondos documentales de valor histórico o BIC-CDA?', SNP,
                     'Perfiles, asignación de funciones y soportes de competencias.',
                     'Personal especializado para fondos históricos o BIC-CDA',
                     hijas=[multi('modalidad', 'Modalidad de atención',
                                  ['Personal de planta especializado', 'Contratación externa por prestación de servicios',
                                   'Personal técnico de archivo general sin especialización', 'No cuenta con personal asignado']),
                            texto('perfiles', 'Perfiles (conservación, restauración, archivística histórica o paleografía)')]),
            pregunta(63, '¿Los acervos patrimoniales o BIC-CDA de la entidad se encuentran organizados y descritos mediante instrumentos y estándares de descripción archivística aplicables?', SNP_NA,
                     'Inventarios normalizados, guías, catálogos, instrumentos descriptivos.',
                     'Descripción normalizada de acervos patrimoniales',
                     hijas=[sel('cobertura', 'Cobertura de descripción', ['Total', 'Parcial', 'Iniciada', 'Solo FUID ordinario o sin inventariar']),
                            texto('estandares', 'Estándares utilizados (ISAD(G), ISDIAH, ISAAR(CPF) u otros)', max_=250),
                            just_na()]),
            pregunta(64, '¿La entidad aplica acciones periódicas de conservación preventiva, restauración o intervención técnica sobre sus acervos patrimoniales o BIC-CDA?', SNP,
                     'Plan de conservación, diagnósticos, registros de intervención, bitácoras ambientales e inspecciones.',
                     'Conservación y restauración de acervos patrimoniales',
                     hijas=[multi('acciones', 'Acciones implementadas',
                                  ['Plan formalizado con intervenciones periódicas', 'Limpieza y monitoreo ambiental puntuales',
                                   'Diagnóstico de deterioro sin intervenciones', 'Ninguna acción']),
                            sel('monitoreo', '¿Se miden temperatura, humedad relativa y plagas?',
                                ['Sí, con equipos y bitácoras', 'Parcialmente', 'No'])]),
            pregunta(65, '¿La entidad tiene formalizado y aplica el procedimiento de transferencia secundaria de fondos de conservación permanente al Archivo Histórico de Medellín o al archivo histórico territorial competente, cuando corresponda?', SNP,
                     'Procedimiento, cronograma, inventarios, actas de transferencia.',
                     'Transferencia secundaria al Archivo Histórico de Medellín',
                     hijas=[sel('estado', 'Estado de las transferencias',
                                ['Cronograma y transferencias ejecutadas', 'Procedimiento formalizado con rezago',
                                 'Fondos históricos en custodia sin transferencia', 'No se han programado transferencias'])]),
            pregunta(66, '¿La entidad dispone de copias digitales de seguridad y estrategias de preservación digital a largo plazo para sus documentos históricos o BIC-CDA, cuando corresponda?', SNP,
                     'Inventarios de copias, especificaciones técnicas, repositorios, procedimientos de consulta y plan de preservación digital.',
                     'Copias digitales y preservación de documentos históricos',
                     hijas=[sel('digitalizacion', 'Estado de digitalización',
                                ['100 % con repositorio seguro', 'Digitalizados parcialmente', 'Proyecto de digitalización en curso', 'No digitalizados']),
                            sel('consulta', '¿La consulta se realiza prioritariamente sobre copias digitales para reducir la manipulación del original?',
                                ['Sí', 'Parcialmente', 'No'])]),
            pregunta(67, '¿Los acervos patrimoniales custodiados incorporan archivos vinculados con la memoria del conflicto armado, los derechos humanos o procesos de interés histórico local del Distrito de Medellín, y se gestionan con los controles de acceso y protección aplicables?', SNP,
                     'Inventarios temáticos, protocolos de consulta, clasificación de acceso y medidas de no revictimización.',
                     'Archivos de memoria, conflicto armado y derechos humanos',
                     hijas=[sel('identificacion', 'Estado de identificación', ['Identificados', 'En proceso de identificación', 'No identificados']),
                            sel('protocolo', 'Protocolo aplicado', ['Protocolo institucional específico de DD. HH. y memoria',
                                                                    'Aplicación de reglas generales de acceso', 'En formulación', 'Sin protocolo'])]),
         ]},
        {'nombre': 'Participación ciudadana',
         'preguntas': [
            pregunta(68, '¿La entidad desarrolla programas de divulgación, proyección social o pedagogía orientados a la visibilización y puesta en valor de sus acervos documentales históricos y BIC-CDA ante la ciudadanía de Medellín?', SNP_NA,
                     'Programas, cronogramas, publicaciones, registros de actividades y medidas de acceso.',
                     'Divulgación y pedagogía de los acervos históricos',
                     hijas=[multi('actividades', 'Tipo y frecuencia de actividades',
                                  ['Exposiciones físicas o virtuales', 'Publicaciones', 'Guías pedagógicas', 'Talleres de memoria',
                                   'Actividades esporádicas en fechas conmemorativas', 'Ninguna']),
                            just_na()]),
            pregunta(69, '¿La entidad incorpora necesidades de la ciudadanía y de los grupos de valor en sus acciones de acceso y divulgación documental, articuladas con las políticas institucionales aplicables?', SNP,
                     'Mecanismos de participación o retroalimentación.', 'Necesidades de la ciudadanía en el acceso y la divulgación',
                     hijas=[justificacion(('no', 'si', 'parcial'), requerido=False)]),
            pregunta(70, '¿La entidad promueve prácticas de acceso inclusivo a la información y al patrimonio documental para distintos grupos de valor?', SNP,
                     'Canales y medidas de accesibilidad e inclusión verificables.', 'Acceso inclusivo a la información y al patrimonio documental',
                     hijas=[justificacion(('no', 'si', 'parcial'), requerido=False)]),
         ]},
        {'nombre': 'Protección del ambiente',
         'preguntas': [
            pregunta(71, '¿La entidad aplica medidas para reducir el impacto ambiental asociado a la producción y gestión de documentos?', SNP,
                     'Acciones verificables de uso racional de papel y otros recursos.', 'Reducción del impacto ambiental de la gestión documental',
                     hijas=[justificacion(('no', 'si', 'parcial'), requerido=False)]),
            pregunta(72, '¿La entidad realiza seguimiento a medidas de uso eficiente de papel, insumos y recursos asociados a la gestión documental?', SNP,
                     'Indicadores o registros de consumo y acciones de mejora.', 'Seguimiento al uso eficiente de papel e insumos',
                     hijas=[justificacion(('no', 'si', 'parcial'), requerido=False)]),
         ]},
    ]},
]


# ----- Índices y utilidades ----------------------------------------------------------
def _indexar():
    preguntas = []
    for comp in COMPONENTES:
        for sub in comp['subcomponentes']:
            for p in sub['preguntas']:
                p['componente'] = comp['clave']
                p['subcomponente'] = sub['nombre']
                for h in p['hijas']:
                    h['campo'] = f"{p['campo']}_{h['id']}"
                    h['show'] = ','.join(f"{p['campo']}={c}" for c in h['cuando'])
                preguntas.append(p)
    return preguntas


PREGUNTAS = _indexar()
POR_CAMPO = {p['campo']: p for p in PREGUNTAS}
TOTAL_PREGUNTAS = len(PREGUNTAS)
assert [p['n'] for p in PREGUNTAS] == list(range(1, 73)), 'La numeración de las preguntas debe ir de 1 a 72'


def es_v04(respuestas):
    return isinstance(respuestas, dict) and (respuestas.get('version') == VERSION or 'p1' in respuestas)


def etiqueta_opcion(p, codigo):
    return dict(p['opciones']).get(codigo, codigo or '')


def puntaje_pregunta(p, respuestas):
    """0, 0,5 o 1; None si no aplica, no se respondió o se excluye."""
    codigo = respuestas.get(p['campo'], '')
    if p['excluir']:
        hija, valor = p['excluir']
        if respuestas.get(f"{p['campo']}_{hija}") == valor:
            return None
    return PUNTOS.get(codigo)


def puntajes_componentes(respuestas):
    resultado = {}
    for comp in COMPONENTES:
        valores = [v for p in PREGUNTAS if p['componente'] == comp['clave']
                   for v in [puntaje_pregunta(p, respuestas)] if v is not None]
        resultado[comp['clave']] = (sum(valores) / len(valores)) if valores else None
    return resultado


def hija_visible(h, respuestas, p):
    return respuestas.get(p['campo'], '') in h['cuando']


def campos_formulario():
    """Nombres de todos los campos del cuestionario, en orden (para Excel)."""
    campos = []
    for p in PREGUNTAS:
        campos.append((p['campo'], f"P{p['n']}. {p['tema']}"))
        for h in p['hijas']:
            if h['tipo'] == 'matriz':
                for i, fila in enumerate(h['filas']):
                    campos.append((f"{h['campo']}_{i}", f"P{p['n']} · {fila}"))
            else:
                campos.append((h['campo'], f"P{p['n']} · {h['etiqueta']}"))
                if h['tipo'] == 'url':
                    campos.append((f"{h['campo']}_nd", f"P{p['n']} · Enlace no disponible"))
                if h.get('otro'):
                    campos.append((f"{h['campo']}_otro", f"P{p['n']} · {h['otro']} (¿cuál?)"))
    return campos
