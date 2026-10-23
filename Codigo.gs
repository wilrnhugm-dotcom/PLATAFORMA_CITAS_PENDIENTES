
/**
 * ==========================================================
 * UGM - GESTIÓN DE SOBRECUPOS
 * NUEVO HOSPITAL ROSALES
 * VERSIÓN 3.0 - AGENDA EXCEL
 * ==========================================================
 *
 * BASE EXISTENTE: UGM2_BASE_ID
 *
 * HOJAS:
 * SOBRECUPOS
 * DISPONIBILIDAD
 * AUDITORIA
 * PERSONAL
 * AGENDA_EXCEL (nueva, se crea automáticamente)
 * PACIENTES (base de pacientes, se crea al cargarla)
 * UGM_CONFIG (oculta, datos de la base de pacientes)
 *
 * REGLAS:
 * - Cupos de sobrecupo por especialidad y día de fin de
 *   semana = Limite de DISPONIBILIDAD (10; Cuidados
 *   Paliativos 5; Nutrición-PE 7) + 1 cupo de URGENCIA
 *   (EXTRA_URGENCIA). El cupo de urgencia solo se usa
 *   cuando los regulares están llenos, exige motivo y
 *   siempre llega como urgente a supervisión, que valida
 *   el caso. Al llenarse se bloquea por restricción médica. Supervisión puede cambiar el
 *   límite general o el de una fecha (LIMITES_FECHA).
 * - No se agendan sobrecupos para el mismo día, salvo
 *   solicitud médica (igual que entre semana).
 * - Toda solicitud entra como PENDIENTE (sin verificar) a
 *   la bandeja de supervisión. Procesarla no libera cupo;
 *   solo rechazar, anular o liberar lo hace.
 * - Redondeo hacia abajo.
 * - Sábado y domingo exclusivamente.
 * - Máximos definidos por especialidad y día.
 * - Jornada completa, sin separar AM y PM.
 * - Las citas pendientes no generan cupos.
 * - Un paciente activo + pendiente cuenta una vez.
 * - Los registros pendientes reservan cupos.
 * - Confirmar no descuenta otra vez.
 * - Rechazar o anular libera el cupo.
 * - Comprobante requerido para confirmar.
 * - Validación bajo LockService.
 * - Sincronización de pantallas cada 10 segundos.
 * - Agenda vigente por 30 minutos para nuevas solicitudes.
 * - Fechas bloqueadas (FECHAS_BLOQUEADAS): no admiten
 *   nuevas programaciones ni asignación de cupos.
 * - DUI: exactamente nueve dígitos (00000000-0). Los DUI
 *   de ocho dígitos de la base de pacientes se completan
 *   con un cero inicial.
 *
 * IMPORTANTE:
 * Piloto. Las claves compartidas no constituyen
 * autenticación institucional para datos clínicos reales.
 * ==========================================================
 */

const UGM3 = {
  // Debe coincidir con la versión de Tecnico, Supervisor,
  // Estilos y Comun: así la pantalla avisa si falta
  // actualizar algún archivo.
  VERSION: '2026.10.23-1',

  TZ: 'America/El_Salvador',

  // Claves exactas (mayúsculas, sin espacios).
  ROLES: {
    SUPERVISOR: 'SUP1',
    AUTORIZADO: 'HNRUGM$',
    TECNICO: 'UGM2026',
    // Unidad de Atención al Paciente: solo consulta por DUI.
    UAP: 'UAP2026'
  },

  // Sin nuevas programaciones ni asignación de cupos.
  FECHAS_BLOQUEADAS: [
    '2026-12-24', '2026-12-25', '2026-12-26',
    '2026-12-27', '2026-12-31',
    '2027-01-01', '2027-01-02', '2027-01-03'
  ],

  // Citas que NO se toman en cuenta (no son especialidad):
  // no se cargan de la agenda ni de la base, ni aparecen en
  // Consultas del día, asistencia o reportes. Se compara
  // sin tildes ni mayúsculas y por el inicio del nombre.
  EXCLUIR_ESPECIALIDADES: ['MEDICINA GENERAL'],

  // Cupos extra por especialidad y día, por encima del
  // límite regular, solo para casos urgentes validados
  // por supervisión (p. ej. 10 regulares + 1 = 11).
  EXTRA_URGENCIA: 1,
  MIN_MOTIVO_URGENCIA: 15,

  // Especialidades del programa extraordinario de los
  // SÁBADOS. Los DOMINGOS todo es extraordinario.
  ESPECIALIDADES_SABADO: [
    'Cardiología','Cirugía General','Cirugía Vascular Periférica',
    'Cuidados Paliativos','Endocrinología','Fisiatría',
    'Gastroenterología','Ginecología','Medicina Familiar',
    'Medicina Familiar - Diabetología','Medicina Interna',
    'Nefrología','Neurología','Nutriología','Oftalmología',
    'Neurocirugía','Ortopedia y Traumatología',
    'Otorrinolaringología','Urología'
  ],

  VIGENCIA_MINUTOS: 30,
  ALERTA_MINUTOS: 60,
  MAX_CITAS: 15000,
  MAX_ARCHIVO_MB: 5,
  MAX_PACIENTES: 250000,
  LOTE_PACIENTES: 5000,
  PAGINA_PACIENTES: 20000,
  MAX_RESULTADOS: 20,

  // Solicitudes sin verificar se consideran urgentes si
  // la cita es en pocos días o llevan horas esperando.
  URGENTE_DIAS: 2,
  URGENTE_HORAS: 2,

  SABADO: {
    'Medicina Familiar - Diabetología': 10,
    'Gastroenterología': 10,
    'Medicina Familiar': 10,
    'Cuidados Paliativos': 5,
    'Nefrología': 10,
    'Cirugía Vascular Periférica': 10,
    'Fisiatría': 10,
    'Endocrinología': 10,
    'Urología': 10,
    'Neurología': 10,
    'Neurocirugía': 10,
    'Cirugía General': 10,
    'Medicina Interna': 10,
    'Nutriología': 10,
    'Nutriología-Diabetología': 10,
    'Otorrinolaringología': 10,
    'Cardiología - Estudios Especiales': null,
    'Ortopedia y Traumatología': 10,
    'Cardiología': 10,
    'Nutrición-PE': 7,
    'Ginecología': 10,
    'Oftalmología': 10
  },

  DOMINGO: {
    'Gastroenterología': 10,
    'Endocrinología': 10,
    'Urología': 10,
    'Neurología': 10,
    'Neurocirugía': 10,
    'Cirugía General': 10,
    'Medicina Interna': 10,
    'Ortopedia y Traumatología': 10,
    'Cardiología': 10,
    'Oftalmología': 10
  },

  HEADERS: {
    SOBRECUPOS: [
      'ID',
      'FechaSobrecupo',
      'Especialidad',
      'Nombres',
      'Documento',
      'SolicitadoPor',
      'RegistradoPor',
      'Sesion',
      'FechaRegistro',
      'Estado',
      'GestionadoPor',
      'FechaGestion',
      'ComprobanteID',
      'Observacion',
      'Tipo',
      'TicketFrente',
      'TicketReverso',
      'Urgencia',
      'MotivoUrgencia'
    ],

    DISPONIBILIDAD: [
      'Dia', 'Especialidad', 'Limite'
    ],

    AUDITORIA: [
      'FechaHora', 'Usuario', 'Perfil',
      'Accion', 'Referencia', 'Detalle'
    ],

    PERSONAL: [
      'Sesion', 'Nombre', 'Perfil',
      'Ingreso', 'UltimaActividad',
      'Actividad', 'Estado'
    ],

    AGENDA_EXCEL: [
      'CargaID',
      'Fecha',
      'Especialidad',
      'Estado',
      'Documento',
      'Paciente',
      'IDExterno',
      'FechaCreacion',
      'Expediente'
    ],

    LIMITES_FECHA: [
      'Fecha',
      'Especialidad',
      'Limite',
      'Usuario',
      'FechaHora'
    ],

    ASISTENCIA: [
      'Clave',
      'Fecha',
      'Especialidad',
      'Documento',
      'Paciente',
      'Origen',
      'Referencia',
      'Asistencia',
      'MarcadoPor',
      'Sesion',
      'FechaHora'
    ],

    PACIENTES: [
      'Nombre',
      'Documento',
      'DocumentoOriginal',
      'Expediente',
      'Especialidad',
      'Busqueda',
      'Fecha',
      'Hora',
      'Aumento',
      'ObservacionUCC',
      'Correlativo'
    ]
  }
};

/* ==========================================================
   APLICACIÓN
   ========================================================== */

// ?vista=supervisor abre Supervisor.html; cualquier
// otro enlace abre Tecnico.html.
function doGet(e) {
  const vista = texto3_(
    e && e.parameter && e.parameter.vista
  ).toLowerCase();

  const pantalla =
    vista === 'supervisor' ? ['Supervisor','UGM | Supervisión'] :
    vista === 'uap' ? ['Uap','UGM | Atención al Paciente'] :
    ['Tecnico','UGM | Técnicos'];

  return HtmlService
    .createTemplateFromFile(pantalla[0])
    .evaluate()
    .setTitle(pantalla[1])
    .addMetaTag(
      'viewport',
      'width=device-width,initial-scale=1,viewport-fit=cover'
    );
}

// Inserta Estilos.html y Comun.html en cada plataforma.
function include(nombre) {
  return HtmlService
    .createHtmlOutputFromFile(nombre)
    .getContent();
}

/* ==========================================================
   UTILIDADES
   ========================================================== */

function props3_() {
  return PropertiesService.getScriptProperties();
}

function ahora3_() {
  return new Date().toISOString();
}

function texto3_(v) {
  return String(v == null ? '' : v).trim();
}

function norm3_(v) {
  return texto3_(v)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s*[-–—]\s*/g, '-')
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

function fecha3_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(
      v, UGM3.TZ, 'yyyy-MM-dd'
    );
  }

  const s = texto3_(v);

  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    throw Error('Fecha inválida.');
  }

  const d = new Date(s + 'T12:00:00Z');

  if (isNaN(d.getTime()) ||
      d.toISOString().slice(0,10) !== s) {
    throw Error('Fecha inválida.');
  }

  return s;
}

function dia3_(fecha) {
  const n = new Date(
    fecha3_(fecha) + 'T12:00:00Z'
  ).getUTCDay();

  return n === 6 ? 'SABADO' :
         n === 0 ? 'DOMINGO' : 'SEMANA';
}

function bloqueada3_(fecha) {
  return UGM3.FECHAS_BLOQUEADAS.includes(fecha);
}

// Devuelve el DUI como 00000000-0, o '' si no tiene
// nueve dígitos. Ocho dígitos: se agrega un cero inicial.
function dui3_(valor) {
  const t = texto3_(valor).replace(/\.0+$/, '');
  const digitos = t.replace(/\D/g, '');

  // Excel guarda el DUI como número y pierde los ceros
  // iniciales (00916089-9 -> 9160899, 000969393 -> 969393):
  // se completan hasta 9 dígitos.
  const numerico = /^\d{5,8}$/.test(t);

  const nueve = numerico
    ? digitos.padStart(9, '0')
    : digitos.length === 8 ? '0' + digitos : digitos;

  return nueve.length === 9
    ? nueve.slice(0,8) + '-' + nueve.slice(8)
    : '';
}

function busqueda3_(v) {
  return norm3_(v)
    .replace(/[^A-Z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function seguro3_(v) {
  const s = texto3_(v);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function fallo3_(codigo, mensaje) {
  throw Error(codigo + '|' + mensaje);
}

/* ==========================================================
   BASE DE GOOGLE SHEETS EXISTENTE
   ========================================================== */

// Se abre una sola vez por ejecución (más rápido).
let libro3_ = null;

function base3_() {
  if (libro3_) return libro3_;

  const id = props3_().getProperty('UGM2_BASE_ID');

  if (!id) {
    throw Error(
      'No está vinculada la base UGM2_BASE_ID.'
    );
  }

  try {
    libro3_ = SpreadsheetApp.openById(id);
  } catch(e) {
    throw Error(
      'La clave es correcta, pero el sistema no puede '+
      'abrir la base de Google Sheets. En Apps Script: '+
      'Implementar > Gestionar implementaciones > '+
      'Ejecutar como "Yo", con una cuenta que tenga '+
      'acceso de edición a la hoja. Ejecuta '+
      'diagnosticoUGM() para revisar. Detalle: '+
      e.message
    );
  }

  return libro3_;
}

/**
 * Crea en el libro las hojas que falten, con sus
 * encabezados y los límites por especialidad. No borra
 * ni modifica hojas existentes.
 */
function prepararHojas3_(libro) {
  const creadas = [];

  const crear = (nombre, cabeceras) => {
    if (libro.getSheetByName(nombre)) return null;

    // Reutiliza la hoja vacía inicial del libro nuevo.
    const vacia = libro.getSheets().find(h =>
      h.getLastRow() === 0 &&
      !Object.keys(UGM3.HEADERS).includes(h.getName())
    );

    const sh = vacia || libro.insertSheet(nombre);
    sh.setName(nombre);
    sh.getRange(1,1,1,cabeceras.length)
      .setValues([cabeceras])
      .setBackground('#10456B')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    sh.setFrozenRows(1);
    creadas.push(nombre);
    return sh;
  };

  const sobrecupos = crear(
    'SOBRECUPOS', UGM3.HEADERS.SOBRECUPOS
  );

  if (sobrecupos) {
    sobrecupos.getRange('B:B').setNumberFormat('@');
    sobrecupos.getRange('E:E').setNumberFormat('@');
  }

  const disponibilidad = crear(
    'DISPONIBILIDAD', UGM3.HEADERS.DISPONIBILIDAD
  );

  if (disponibilidad) {
    const limites = [];

    [['SABADO', UGM3.SABADO], ['DOMINGO', UGM3.DOMINGO]]
      .forEach(([dia, tabla]) => {
        Object.keys(tabla).forEach(esp => {
          limites.push([
            dia, esp,
            tabla[esp] === null ? 'N/A' : tabla[esp]
          ]);
        });
      });

    disponibilidad.getRange(2,1,limites.length,3)
      .setValues(limites);
  }

  crear('AUDITORIA', UGM3.HEADERS.AUDITORIA);
  crear('PERSONAL', UGM3.HEADERS.PERSONAL);

  const asistencia = crear(
    'ASISTENCIA', UGM3.HEADERS.ASISTENCIA
  );

  if (asistencia) {
    asistencia.getRange('A:G').setNumberFormat('@');
  }

  const limites = crear(
    'LIMITES_FECHA', UGM3.HEADERS.LIMITES_FECHA
  );

  if (limites) limites.getRange('A:B').setNumberFormat('@');

  const sc = libro.getSheetByName('SOBRECUPOS');
  if (sc) asegurarColumnas3_(sc, UGM3.HEADERS.SOBRECUPOS);

  return creadas;
}

/**
 * INSTALACIÓN DE LA PLATAFORMA NUEVA.
 * 1. Crear una hoja de cálculo nueva.
 * 2. Extensiones > Apps Script, pegar Código.gs e Index.
 * 3. Elegir instalarUGM y pulsar Ejecutar (una vez).
 * Esa hoja de cálculo queda como base de datos.
 */
function instalarUGM() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();

  if (!libro) {
    throw Error(
      'Abre Apps Script desde la hoja de cálculo '+
      '(Extensiones > Apps Script) y vuelve a ejecutar. '+
      'Si el proyecto no está dentro de una hoja, '+
      'usa crearBaseUGM().'
    );
  }

  const creadas = prepararHojas3_(libro);
  const p = props3_();

  p.setProperty('UGM2_BASE_ID', libro.getId());

  if (!p.getProperty('UGM2_SECRET')) {
    p.setProperty(
      'UGM2_SECRET',
      Utilities.getUuid() + Utilities.getUuid()
    );
  }

  libro3_ = null;
  garantizarAgenda3_();

  const texto =
    'Plataforma instalada. Base de datos: ' +
    libro.getName() + ' (' + libro.getUrl() + ')' +
    '\nHojas creadas: ' +
    (creadas.length ? creadas.join(', ') : 'ninguna (ya existían)') +
    '\nSiguiente paso: Implementar > Nueva implementación' +
    ' > Aplicación web > Ejecutar como "Yo".';

  console.log(texto);
  return texto;
}

/**
 * Para proyectos que NO están dentro de una hoja:
 * crea una base nueva en la cuenta que ejecuta y la
 * vincula. Si la base actual abre, no hace cambios.
 */
function crearBaseUGM() {
  const p = props3_();
  const actual = p.getProperty('UGM2_BASE_ID');

  if (actual) {
    let abre = false;
    try {
      SpreadsheetApp.openById(actual);
      abre = true;
    } catch(e) {}

    if (abre) {
      throw Error(
        'La base actual abre correctamente; '+
        'no se creó otra.'
      );
    }
  }

  const libro = SpreadsheetApp.create(
    'UGM 2026 - Base de sobrecupos'
  );

  prepararHojas3_(libro);

  p.setProperty('UGM2_BASE_ID', libro.getId());
  libro3_ = null;
  garantizarAgenda3_();

  const texto = 'Base creada y vinculada: ' + libro.getUrl();
  console.log(texto);
  return texto;
}

/**
 * Ejecutar desde el editor de Apps Script (botón
 * Ejecutar) para revisar la conexión con la base.
 * También solicita los permisos si faltan.
 */
function diagnosticoUGM() {
  const r = [];
  const id = props3_().getProperty('UGM2_BASE_ID');

  try {
    r.push('Cuenta que ejecuta: ' +
      Session.getEffectiveUser().getEmail());
  } catch(e) {
    r.push('Cuenta que ejecuta: no disponible');
  }

  r.push('UGM2_BASE_ID: ' + (id || 'NO CONFIGURADO'));

  try {
    const libro = SpreadsheetApp.openById(id);
    r.push('Base: OK -> ' + libro.getName());

    ['SOBRECUPOS','DISPONIBILIDAD',
     'AUDITORIA','PERSONAL'].forEach(n => {
      r.push('Hoja ' + n + ': ' +
        (libro.getSheetByName(n) ? 'OK' : 'FALTA'));
    });
  } catch(e) {
    r.push('Base: ERROR -> ' + e.message);
  }

  r.push('Clave supervisor: "' +
    UGM3.ROLES.SUPERVISOR + '"');
  r.push('Clave técnico: "' +
    UGM3.ROLES.TECNICO + '"');

  const texto = r.join('\n');
  console.log(texto);
  return texto;
}

const revisadas3_ = {};

// Completa encabezados nuevos sin mover los existentes.
function asegurarColumnas3_(sh, cabeceras) {
  const n = cabeceras.length;

  if (sh.getMaxColumns() < n) {
    sh.insertColumnsAfter(
      sh.getMaxColumns(), n - sh.getMaxColumns()
    );
  }

  const actual = sh.getRange(1,1,1,n).getValues()[0];

  if (actual.every((v,i) => String(v) === cabeceras[i])) {
    return;
  }

  sh.getRange(1,1,1,n).setValues([
    cabeceras.map((c,i) => actual[i] === '' ? c : actual[i])
  ]);
}

function hoja3_(nombre) {
  const libro = base3_();
  let sh = libro.getSheetByName(nombre);

  // Hojas nuevas de esta versión se crean solas.
  if (!sh && ['ASISTENCIA','LIMITES_FECHA'].includes(nombre)) {
    prepararHojas3_(libro);
    sh = libro.getSheetByName(nombre);
  }

  if (!sh) {
    throw Error('No existe la hoja: ' + nombre);
  }

  if (!revisadas3_[nombre] && UGM3.HEADERS[nombre] &&
      ['SOBRECUPOS','ASISTENCIA'].includes(nombre)) {
    asegurarColumnas3_(sh, UGM3.HEADERS[nombre]);
    revisadas3_[nombre] = true;
  }

  return sh;
}

function hoy3_() {
  return Utilities.formatDate(
    new Date(), UGM3.TZ, 'yyyy-MM-dd'
  );
}

function columna3_(hoja, campo) {
  return UGM3.HEADERS[hoja].indexOf(campo);
}

function garantizarAgenda3_() {
  const libro = base3_();

  let sh = libro.getSheetByName('AGENDA_EXCEL');

  if (!sh) {
    sh = libro.insertSheet('AGENDA_EXCEL');

    sh.getRange(
      1, 1, 1,
      UGM3.HEADERS.AGENDA_EXCEL.length
    ).setValues([
      UGM3.HEADERS.AGENDA_EXCEL
    ]);

    sh.getRange(1,1,1,UGM3.HEADERS.AGENDA_EXCEL.length)
      .setBackground('#10456B')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');

    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('E:E').setNumberFormat('@');
  }

  // Agrega columnas nuevas (FechaCreacion) a la hoja que ya
  // existe, sin mover las anteriores.
  if (!revisadas3_.AGENDA_EXCEL) {
    const cab = UGM3.HEADERS.AGENDA_EXCEL;
    const n = cab.length;
    if (sh.getMaxColumns() < n ||
        String(sh.getRange(1,n).getValue()) !== cab[n-1]) {
      asegurarColumnas3_(sh, cab);
      sh.getRange(1,n,sh.getMaxRows(),1).setNumberFormat('@');
    }
    revisadas3_.AGENDA_EXCEL = true;
  }

  return sh;
}

function prepararSistemaUGM() {
  garantizarAgenda3_();

  return {
    correcto: true,
    base: base3_().getUrl(),
    mensaje:
      'Agenda Excel preparada sin eliminar registros.'
  };
}

function leer3_(nombre) {
  const sh = nombre === 'AGENDA_EXCEL'
    ? garantizarAgenda3_()
    : hoja3_(nombre);

  const ultima = sh.getLastRow();

  if (ultima < 2) return [];

  const matriz = sh.getRange(
    1, 1, ultima,
    UGM3.HEADERS[nombre].length
  ).getValues();

  const cabeceras = matriz.shift().map(String);

  return matriz
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(fila => {
      const obj = {};

      cabeceras.forEach((campo,i) => {
        let v = fila[i];

        if (v instanceof Date) {
          if ([
            'Fecha',
            'FechaSobrecupo'
          ].includes(campo)) {
            v = Utilities.formatDate(
              v, UGM3.TZ, 'yyyy-MM-dd'
            );
          } else {
            v = Utilities.formatDate(
              v, UGM3.TZ,
              "yyyy-MM-dd'T'HH:mm:ss"
            );
          }
        }

        obj[campo] = v;
      });

      return obj;
    });
}

function insertar3_(nombre, datos) {
  const sh = hoja3_(nombre);

  sh.appendRow(
    UGM3.HEADERS[nombre].map(campo => {
      const valor = datos[campo] === undefined
        ? '' : datos[campo];

      return typeof valor === 'string'
        ? seguro3_(valor)
        : valor;
    })
  );
}

function auditar3_(usuario, perfil, accion, ref, detalle) {
  insertar3_('AUDITORIA', {
    FechaHora: ahora3_(),
    Usuario: usuario,
    Perfil: perfil,
    Accion: accion,
    Referencia: ref || '',
    Detalle: texto3_(detalle).slice(0,350)
  });
}

/* ==========================================================
   CATÁLOGO Y NORMALIZACIÓN
   ========================================================== */

function catalogo3_() {
  return [...new Set([
    ...Object.keys(UGM3.SABADO),
    ...Object.keys(UGM3.DOMINGO)
  ])];
}

function especialidad3_(valor) {
  const normal = norm3_(valor);

  const alias = {
    'NUTRICION-PE': 'Nutrición-PE',
    'NUTRICION - PE': 'Nutrición-PE',
    'NUTRIOLOGIA-DIABETOLOGIA':
      'Nutriología-Diabetología',
    'ORTOPEDIA Y TRAUMATOLOGIA':
      'Ortopedia y Traumatología',
    'MEDICINA FAMILIAR-DIABETOLOGIA':
      'Medicina Familiar - Diabetología'
  };

  if (alias[normal]) return alias[normal];

  return catalogo3_().find(
    e => norm3_(e) === normal
  ) || null;
}

function estadoCita3_(valor) {
  const s = norm3_(valor);

  if ([
    'ACTIVA',
    'ACTIVO',
    'CONFIRMADA',
    'CONFIRMADO'
  ].includes(s)) return 'ACTIVA';

  if (s.startsWith('PENDIENTE')) {
    return 'PENDIENTE';
  }

  if (s.startsWith('CANCELAD') ||
      s.startsWith('ANULAD')) {
    return 'CANCELADA';
  }

  if (s.startsWith('PROCESAD')) {
    return 'PROCESADA';
  }

  return 'OTRA';
}

function documento3_(valor) {
  return texto3_(valor).replace(/\s+/g,'');
}

function fechaAgenda3_(valor) {
  if (valor instanceof Date) {
    return fecha3_(valor);
  }

  const s = texto3_(valor);

  let m = s.match(
    /^(\d{4})-(\d{1,2})-(\d{1,2})/
  );

  if (m) {
    return fecha3_(
      m[1]+'-'+m[2].padStart(2,'0')+
      '-'+m[3].padStart(2,'0')
    );
  }

  m = s.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})/
  );

  if (m) {
    return fecha3_(
      m[3]+'-'+m[2].padStart(2,'0')+
      '-'+m[1].padStart(2,'0')
    );
  }

  throw Error('Fecha de agenda no reconocida.');
}

/* ==========================================================
   SESIONES
   ========================================================== */

function firmar3_(sesion) {
  const cuerpo = [
    sesion.id,
    sesion.nombre,
    sesion.perfil,
    sesion.inicio
  ].join('|');

  const secreto = props3_().getProperty(
    'UGM2_SECRET'
  );

  if (!secreto) {
    throw Error('Falta inicializar las sesiones.');
  }

  const firma = Utilities
    .computeHmacSha256Signature(cuerpo,secreto)
    .map(b =>
      ('0'+(b & 255).toString(16)).slice(-2)
    ).join('');

  // UTF-8 explícito: nombres con tildes o ñ (José, Peña)
  // deben producir el mismo token al validarlo.
  return Utilities.base64EncodeWebSafe(
    cuerpo+'|'+firma, Utilities.Charset.UTF_8
  );
}

// UAP solo puede usar su consulta por DUI (permitirUAP).
function validar3_(token, permitirUAP) {
  if (!token) throw Error('Inicia sesión.');

  let partes;

  try {
    partes = Utilities.newBlob(
      Utilities.base64DecodeWebSafe(token)
    ).getDataAsString('UTF-8').split('|');
  } catch(e) {
    throw Error('Sesión inválida.');
  }

  if (partes.length !== 5) {
    throw Error('Sesión inválida.');
  }

  const s = {
    id: partes[0],
    nombre: partes[1],
    perfil: partes[2],
    inicio: partes[3]
  };

  if (firmar3_(s) !== token) {
    throw Error('Sesión no autorizada.');
  }

  if (Date.now()-Number(s.inicio) > 43200000) {
    throw Error('Sesión vencida.');
  }

  const conectado = leer3_('PERSONAL').some(r =>
    String(r.Sesion) === s.id &&
    String(r.Estado) === 'CONECTADO'
  );

  if (!conectado) throw Error('Sesión cerrada.');

  if (s.perfil === 'UAP' && !permitirUAP) {
    throw Error('El acceso de UAP solo permite la consulta por DUI.');
  }

  return s;
}

// Cada enlace (?vista=tecnico / ?vista=supervisor)
// solo admite sus perfiles. Sin parámetro, admite todos.
const VISTAS3 = {
  tecnico: ['TECNICO','AUTORIZADO'],
  supervisor: ['SUPERVISOR','AUTORIZADO'],
  uap: ['UAP','SUPERVISOR','AUTORIZADO']
};

function entrar(nombre, clave, vista) {
  nombre = texto3_(nombre).slice(0,70);

  if (nombre.length < 3 || nombre.includes('|')) {
    throw Error('Escribe un nombre válido.');
  }

  let perfil = '';

  Object.keys(UGM3.ROLES).forEach(rol => {
    if (UGM3.ROLES[rol] === String(clave||'')) {
      perfil = rol;
    }
  });

  if (!perfil) throw Error('Clave incorrecta.');

  const v = texto3_(vista).toLowerCase();

  // La clave de UAP solo abre su propio enlace.
  if (perfil === 'UAP' && v !== 'uap') {
    throw Error(
      'La clave de UAP solo funciona en el enlace de '+
      'Atención al Paciente (?vista=uap).'
    );
  }

  if (VISTAS3[v] && !VISTAS3[v].includes(perfil)) {
    throw Error(
      v === 'uap'
        ? 'Esta es la pantalla de Atención al Paciente (UAP). '+
          'Usa la clave de UAP.'
        : v === 'tecnico'
        ? 'Esta es la plataforma de técnicos. '+
          'Ingresa por el enlace de supervisión.'
        : 'Esta es la plataforma de supervisión. '+
          'Ingresa por el enlace de técnicos.'
    );
  }

  base3_();

  if (!props3_().getProperty('UGM2_SECRET')) {
    props3_().setProperty(
      'UGM2_SECRET',
      Utilities.getUuid()+Utilities.getUuid()
    );
  }

  const sesion = {
    id: Utilities.getUuid(),
    nombre,
    perfil,
    inicio: String(Date.now())
  };

  const token = firmar3_(sesion);

  insertar3_('PERSONAL',{
    Sesion: sesion.id,
    Nombre: nombre,
    Perfil: perfil,
    Ingreso: ahora3_(),
    UltimaActividad: ahora3_(),
    Actividad: 'Ingreso',
    Estado: 'CONECTADO'
  });

  try {
    auditar3_(
      nombre,perfil,'INICIO_SESION',sesion.id,''
    );
  } catch(e) {
    console.error(e.message);
  }

  return {token,nombre,perfil,sesion:sesion.id,version:UGM3.VERSION};
}

function modificarSesion3_(s,actividad,cerrar) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sh = hoja3_('PERSONAL');
    const matriz = sh.getDataRange().getValues();

    for (let i=1;i<matriz.length;i++) {
      if (String(matriz[i][0]) === s.id) {
        if (cerrar) {
          sh.getRange(i+1,5,1,3).setValues([[
            ahora3_(),
            'Sesión cerrada',
            'DESCONECTADO'
          ]]);
        } else {
          sh.getRange(i+1,5,1,2).setValues([[
            ahora3_(),
            texto3_(actividad).slice(0,100)
          ]]);
        }
        return true;
      }
    }

    return false;
  } finally {
    lock.releaseLock();
  }
}

function salir(token) {
  const s = validar3_(token, true);

  modificarSesion3_(s,'Sesión cerrada',true);

  try {
    auditar3_(
      s.nombre,s.perfil,'CIERRE_SESION',s.id,''
    );
  } catch(e) {
    console.error(e.message);
  }

  return {ok:true};
}

/* ==========================================================
   CARGA DE AGENDA EXCEL
   ========================================================== */

function importarAgenda(token, registros, archivo) {
  const usuario = validar3_(token);

  // Solo supervisión carga la agenda del día.
  if (usuario.perfil !== 'SUPERVISOR') {
    throw Error(
      'Solo supervisión puede cargar la agenda del día.'
    );
  }

  if (!Array.isArray(registros) ||
      registros.length === 0) {
    throw Error('El Excel no contiene registros.');
  }

  if (registros.length > UGM3.MAX_CITAS) {
    throw Error(
      'El Excel supera '+UGM3.MAX_CITAS+
      ' registros permitidos en este piloto.'
    );
  }

  const validos = [];
  const fechas = new Set();
  const porFecha = {};
  const COLS = UGM3.HEADERS.AGENDA_EXCEL.length;
  let conCreacion = 0;
  let conExpediente = 0;

  const errores = {
    sinFecha:0,
    sinEspecialidad:0,
    fueraDeCatalogo:0,
    estadoNoReconocido:0,
    excluidas:0
  };

  registros.forEach((r,i) => {
    let f;

    try {
      f = fechaAgenda3_(r.fecha);
    } catch(e) {
      errores.sinFecha++;
      return;
    }

    // Medicina General (F2) no se toma en cuenta.
    if (excluida3_(r.especialidad)) {
      errores.excluidas++;
      return;
    }

    // Otras especialidades fuera del catálogo de sobrecupos
    // sí se guardan: los técnicos marcan su asistencia igual.
    const catalogo = especialidad3_(r.especialidad);
    const esp = catalogo ||
      texto3_(r.especialidad).replace(/\s+/g,' ').slice(0,100);

    if (!esp) {
      errores.sinEspecialidad++;
      return;
    }

    if (!catalogo) errores.fueraDeCatalogo++;

    const estado = estadoCita3_(r.estado);

    if (estado === 'OTRA') {
      errores.estadoNoReconocido++;
      return;
    }

    const doc = documento3_(r.documento);
    // El reporte trae "EXPEDIENTE - NOMBRE": se deja el nombre.
    const paciente = texto3_(r.paciente)
      .replace(/^\s*\d[\d\-\/]*\s*-\s*/,'')
      .slice(0,150);
    const id = texto3_(r.id).slice(0,100);

    if (!doc && !paciente && !id) {
      return;
    }

    fechas.add(f);
    if (estado !== 'CANCELADA') porFecha[f] = (porFecha[f] || 0) + 1;

    let creacion = '';
    try {
      if (texto3_(r.creacion)) creacion = fechaAgenda3_(r.creacion);
    } catch(e) {}
    if (creacion) conCreacion++;

    // Expediente: es lo que trae el QR del ticket de la cita.
    const expediente = texto3_(r.expediente)
      .replace(/\.0+$/,'').slice(0,30);
    if (expediente) conExpediente++;

    validos.push([
      '', f, esp, estado,
      doc, paciente, id || String(i+1), creacion, expediente
    ]);
  });

  if (!validos.length) {
    throw Error(
      'No se encontraron citas válidas. '+
      'Revisa las columnas de fecha, especialidad y estado.'
    );
  }

  // Evita sustituir la base por un archivo
  // masivamente incompatible.
  if (validos.length <
      (registros.length - errores.excluidas) * 0.5) {
    throw Error(
      'Se rechazó la carga: menos del 50% de las '+
      'filas fueron reconocidas. Verifica el Excel.'
    );
  }

  const cargaID = Utilities.getUuid();

  validos.forEach(r => r[0] = cargaID);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  let anterior = null;

  try {
    const sh = garantizarAgenda3_();
    const ultima = sh.getLastRow();

    if (ultima > 1) {
      anterior = sh.getRange(
        2,1,ultima-1,COLS
      ).getValues();
      sh.getRange(
        2,1,ultima-1,COLS
      ).clearContent();
    }

    try {
      for (let i=0;i<validos.length;i+=500) {
        const lote = validos
          .slice(i,i+500)
          .map(fila => fila.map(seguro3_));

        sh.getRange(
          i+2,1,lote.length,COLS
        ).setValues(lote);
      }

      SpreadsheetApp.flush();

    } catch(error) {
      // Recuperación de la hoja anterior si falla
      // la escritura de la nueva agenda.
      const fin = sh.getLastRow();

      if (fin > 1) {
        sh.getRange(
          2,1,fin-1,COLS
        ).clearContent();
      }

      if (anterior && anterior.length) {
        for (let i=0;i<anterior.length;i+=500) {
          const lote=anterior.slice(i,i+500);
          sh.getRange(
            i+2,1,lote.length,COLS
          ).setValues(lote);
        }
      }

      SpreadsheetApp.flush();
      throw error;
    }

    const p = props3_();
    const hora = ahora3_();

    p.setProperties({
      UGM3_AGENDA_AT:hora,
      UGM3_AGENDA_ID:cargaID,
      UGM3_AGENDA_ARCHIVO:
        texto3_(archivo).slice(0,150),
      UGM3_AGENDA_USUARIO:usuario.nombre,
      UGM3_AGENDA_TOTAL:String(validos.length),
      UGM3_AGENDA_FECHAS:JSON.stringify(porFecha),
      UGM3_AGENDA_CREACION:String(conCreacion),
      UGM3_AGENDA_EXPEDIENTE:String(conExpediente),
      UGM3_ALERTA_AT:''
    });

  } finally {
    lock.releaseLock();
  }

  try {
    auditar3_(
      usuario.nombre,
      usuario.perfil,
      'AGENDA_IMPORTADA',
      cargaID,
      archivo+' · '+validos.length+' citas'
    );
  } catch(e) {
    console.error(e.message);
  }

  return {
    ok:true,
    cargaID,
    total:validos.length,
    fechas:[...fechas].sort(),
    porFecha,
    conCreacion,
    omitidas:registros.length-validos.length-errores.excluidas,
    excluidas:errores.excluidas,
    errores,
    mensaje:'Agenda Excel actualizada correctamente.'
  };
}

/* ==========================================================
   VIGENCIA Y ALERTAS
   ========================================================== */

function estadoAgenda3_() {
  const p = props3_();

  const fecha = p.getProperty(
    'UGM3_AGENDA_AT'
  ) || '';

  const extraordinaria = p.getProperty(
    'UGM3_ALERTA_AT'
  ) || '';

  const minutos = fecha
    ? Math.floor(
        (Date.now()-Date.parse(fecha))/60000
      )
    : null;

  const hayAlerta = Boolean(extraordinaria);

  let porFecha = {};
  try {
    porFecha = JSON.parse(p.getProperty('UGM3_AGENDA_FECHAS') || '{}');
  } catch(e) {}

  return {
    total:Number(p.getProperty('UGM3_AGENDA_TOTAL')) || 0,
    porFecha,
    // Citas con fecha de creación del HIS (para separar
    // extraordinarios de citas previas).
    conCreacion:Number(p.getProperty('UGM3_AGENDA_CREACION')) || 0,
    conExpediente:Number(p.getProperty('UGM3_AGENDA_EXPEDIENTE')) || 0,
    fecha,
    archivo:p.getProperty(
      'UGM3_AGENDA_ARCHIVO'
    ) || '',
    usuario:p.getProperty(
      'UGM3_AGENDA_USUARIO'
    ) || '',
    cargaID:p.getProperty(
      'UGM3_AGENDA_ID'
    ) || '',
    minutos,
    vigente:
      minutos !== null &&
      minutos >= 0 &&
      minutos < UGM3.VIGENCIA_MINUTOS &&
      !hayAlerta,
    // Solo se alerta si no hay agenda o si se pidió una
    // actualización extraordinaria. La antigüedad ya no
    // alerta: así nadie vuelve a subirla sin necesidad.
    alerta:
      minutos === null ||
      hayAlerta,
    extraordinaria:hayAlerta
  };
}

function solicitarActualizacionGlobal(token) {
  const s = validar3_(token);

  if (s.perfil !== 'SUPERVISOR') {
    throw Error('Solo el supervisor puede hacerlo.');
  }

  props3_().setProperty(
    'UGM3_ALERTA_AT',
    ahora3_()
  );

  auditar3_(
    s.nombre,s.perfil,
    'ALERTA_AGENDA','',
    'Actualización extraordinaria solicitada'
  );

  return {ok:true};
}

/* ==========================================================
   PACIENTES ÚNICOS POR FECHA Y ESPECIALIDAD
   ========================================================== */

function identidad3_(r,index) {
  const doc = norm3_(r.Documento);

  if (doc) return 'D:'+doc;

  const paciente = norm3_(r.Paciente);

  if (paciente) return 'N:'+paciente;

  const id = norm3_(r.IDExterno);

  if (id) return 'I:'+id;

  return 'F:'+index;
}

function contarAgenda3_(agenda,fecha,especialidad) {
  const pacientes = new Map();

  agenda.forEach((r,i) => {
    if (texto3_(r.Fecha) !== fecha ||
        norm3_(r.Especialidad) !==
          norm3_(especialidad)) {
      return;
    }

    const identidad = identidad3_(r,i);
    const estado = estadoCita3_(r.Estado);

    if (!pacientes.has(identidad)) {
      pacientes.set(identidad,{
        activa:false,
        pendiente:false
      });
    }

    const item=pacientes.get(identidad);

    if (estado === 'ACTIVA') item.activa=true;
    if (estado === 'PENDIENTE') item.pendiente=true;
  });

  let activas=0;
  let pendientes=0;

  pacientes.forEach(p => {
    if (p.activa) {
      activas++;
    } else if (p.pendiente) {
      pendientes++;
    }
  });

  return {activas,pendientes};
}

function excluida3_(especialidad) {
  const e=norm3_(especialidad);
  return Boolean(e) &&
    UGM3.EXCLUIR_ESPECIALIDADES.some(x=>e.startsWith(x));
}

function numeroLimite3_(v) {
  if (texto3_(v)==='' || norm3_(v)==='N/A') return null;
  const n=Number(v);
  return Number.isInteger(n) && n>=0 ? n : null;
}

// Límites de una fecha: el de la fecha (si supervisión
// lo cambió) o el general del sábado/domingo.
function limitesDeFecha3_(fecha,disponibilidad,excepciones) {
  const dia=dia3_(fecha);
  const mapa=new Map();

  if (dia==='SEMANA') return mapa;

  disponibilidad
    .filter(r=>norm3_(r.Dia)===dia)
    .forEach(r=>{
      mapa.set(norm3_(r.Especialidad),{
        especialidad:texto3_(r.Especialidad),
        limite:numeroLimite3_(r.Limite),
        general:numeroLimite3_(r.Limite),
        porFecha:false
      });
    });

  excepciones
    .filter(r=>texto3_(r.Fecha)===fecha)
    .forEach(r=>{
      const clave=norm3_(r.Especialidad);
      const base=mapa.get(clave);
      mapa.set(clave,{
        especialidad:base ? base.especialidad : texto3_(r.Especialidad),
        limite:numeroLimite3_(r.Limite),
        general:base ? base.general : null,
        porFecha:true,
        cambiadoPor:texto3_(r.Usuario)
      });
    });

  return mapa;
}

function limite3_(fecha,especialidad) {
  const l=limitesDeFecha3_(
    fecha,leer3_('DISPONIBILIDAD'),leer3_('LIMITES_FECHA')
  ).get(norm3_(especialidad));

  return l ? l.limite : null;
}

function ocupados3_(solicitudes,fecha,especialidad) {
  return solicitudes.filter(r =>
    texto3_(r.FechaSobrecupo) === fecha &&
    norm3_(r.Especialidad) ===
      norm3_(especialidad) &&
    !['RECHAZADO','ANULADO'].includes(
      norm3_(r.Estado)
    )
  ).length;
}

function calcular3_(fecha,agenda,solicitudes) {
  if (dia3_(fecha)==='SEMANA') return [];

  const limites=limitesDeFecha3_(
    fecha,leer3_('DISPONIBILIDAD'),leer3_('LIMITES_FECHA')
  );

  const bloqueada=bloqueada3_(fecha);

  return [...limites.values()].map(l => {
    const pacientes=contarAgenda3_(
      agenda,fecha,l.especialidad
    );

    const autorizados=l.limite===null || bloqueada
      ? 0 : l.limite;

    const ocupados=ocupados3_(
      solicitudes,fecha,l.especialidad
    );

    const disponibles=Math.max(0,autorizados-ocupados);

    // Cupo(s) de urgencia: solo cuando los regulares
    // ya están ocupados.
    const urgencia=autorizados>0 ? UGM3.EXTRA_URGENCIA : 0;
    const urgenciaLibre=Math.max(
      0,autorizados+urgencia-Math.max(ocupados,autorizados)
    );

    return {
      especialidad:l.especialidad,
      activas:pacientes.activas,
      pendientes:pacientes.pendientes,
      limite:l.limite,
      general:l.general,
      porFecha:l.porFecha,
      cambiadoPor:l.cambiadoPor || '',
      capacidad:autorizados+urgencia,
      generados:autorizados,
      ocupados,
      disponibles,
      urgencia,
      urgenciaLibre,
      urgenciaUsada:Math.max(0,ocupados-autorizados),
      soloUrgencia:disponibles===0 && urgenciaLibre>0,
      ultimo:disponibles===1,
      completo:autorizados>0 && disponibles===0 &&
        urgenciaLibre===0,
      habilitada:l.limite!==null && !bloqueada,
      bloqueada,
      incidencia:!bloqueada && ocupados>autorizados+urgencia
    };
  });
}

/* ==========================================================
   TABLERO
   ========================================================== */

function tablero(token,fecha) {
  const s=validar3_(token);
  const f=fecha3_(fecha);

  const agenda=leer3_('AGENDA_EXCEL');
  const solicitudes=leer3_('SOBRECUPOS');

  const lista=calcular3_(
    f,agenda,solicitudes
  );

  const visibles=s.perfil==='SUPERVISOR'
    ? solicitudes
    : solicitudes.filter(r =>
        String(r.Sesion) === s.id
      );

  visibles.sort((a,b)=>
    String(b.FechaRegistro||'').localeCompare(
      String(a.FechaRegistro||'')
    )
  );

  const pendientesGlobales=solicitudes
    .filter(r =>
      texto3_(r.FechaSobrecupo)===f &&
      [
        'PENDIENTE',
        'EN GESTION',
        'OBSERVADO'
      ].includes(norm3_(r.Estado))
    ).length;

  const fechas=[...new Set(
    agenda.map(r=>texto3_(r.Fecha))
      .filter(Boolean)
  )].sort();

  const resultado={
    fecha:f,
    dia:dia3_(f),
    bloqueada:bloqueada3_(f),
    fechasBloqueadas:UGM3.FECHAS_BLOQUEADAS,
    agenda:estadoAgenda3_(),
    pacientes:estadoPacientes3_(),
    fechasAgenda:fechas,
    lista,
    solicitudes:visibles.slice(0,200),
    pendientesGlobales,
    personal:[],
    auditoria:[]
  };

  if (s.perfil==='SUPERVISOR') {
    resultado.personal=leer3_('PERSONAL')
      .filter(r =>
        r.Estado==='CONECTADO' &&
        Date.now()-Date.parse(r.UltimaActividad)
          < 90000
      );

    resultado.auditoria=leer3_('AUDITORIA')
      .slice(-70).reverse();
  }

  return resultado;
}

function sincronizar(token,fecha,actividad) {
  const s=validar3_(token);

  // La lectura no espera a modificar PERSONAL.
  const datos=tablero(token,fecha);

  actividad3_(s,actividad);

  return datos;
}

// Registra lo que está haciendo cada usuario (lo ve
// el supervisor en tiempo real). No bloquea la lectura.
function actividad3_(s,actividad) {
  try {
    const lock=LockService.getScriptLock();

    if (lock.tryLock(1000)) {
      try {
        const sh=hoja3_('PERSONAL');
        const m=sh.getDataRange().getValues();

        for(let i=1;i<m.length;i++) {
          if(String(m[i][0])===s.id) {
            sh.getRange(i+1,5,1,2)
              .setValues([[
                ahora3_(),
                texto3_(actividad).slice(0,100)
              ]]);
            break;
          }
        }
      } finally {
        lock.releaseLock();
      }
    }
  } catch(e) {
    console.error(e.message);
  }
}

/* ==========================================================
   GUARDAR SOBRECUPO
   ========================================================== */

const EXTENSIONES3={
  'image/jpeg':'jpg',
  'image/png':'png',
  'image/webp':'webp',
  'application/pdf':'pdf'
};

function ticket3_(archivo) {
  if (!archivo || !archivo.base64) return null;

  const mime=texto3_(archivo.mime);

  return {
    mime,
    bytes:bytesComprobante3_(mime,archivo.base64)
  };
}

/**
 * Solicitud de sobrecupo.
 * - Sábado y domingo: según cupos disponibles.
 * - Otro día: solo como solicitud médica, con ticket
 *   (frente y reverso sellado y firmado por el médico).
 * - Fechas bloqueadas: nunca.
 */
function crearSolicitud(token,datos,archivos) {
  const s=validar3_(token);

  if (!datos || typeof datos!=='object') {
    throw Error('Formulario inválido.');
  }

  [
    'especialidad','nombres','documento',
    'fecha','solicitadoPor'
  ].forEach(k => {
    if (!texto3_(datos[k])) {
      throw Error(
        'Completa los cinco campos obligatorios.'
      );
    }
  });

  const fecha=fecha3_(datos.fecha);
  const especialidad=especialidad3_(
    datos.especialidad
  );
  const documento=texto3_(datos.documento);

  if (!/^[0-9]{8}-[0-9]$/.test(documento)) {
    fallo3_(
      'DUI_INVALIDO',
      'El DUI debe tener exactamente nueve '+
      'dígitos en formato 00000000-0.'
    );
  }

  if (bloqueada3_(fecha)) {
    fallo3_(
      'FECHA_NO_DISPONIBLE',
      'La fecha '+fecha+' no está disponible. '+
      'No se permiten nuevas programaciones ni '+
      'asignación de cupos.'
    );
  }

  if (fecha<hoy3_()) {
    fallo3_(
      'FECHA_NO_DISPONIBLE',
      'No se pueden solicitar sobrecupos para '+
      'fechas pasadas.'
    );
  }

  if (!especialidad) {
    fallo3_(
      'ESPECIALIDAD_NO_HABILITADA',
      'Selecciona una especialidad del catálogo.'
    );
  }

  const finDeSemana=dia3_(fecha)!=='SEMANA';
  const mismoDia=fecha===hoy3_();

  // Entre semana o para el mismo día: solo si el médico
  // lo exige (solicitud médica con ticket).
  const medico=!finDeSemana || mismoDia;

  if (medico && datos.tipo!=='MEDICO') {
    fallo3_(
      'DIA_NO_HABILITADO',
      mismoDia
        ? 'No se agendan sobrecupos para el mismo día. '+
          'Solo si el médico lo exige: marca "Solicitud '+
          'médica" y adjunta el ticket sellado y firmado.'
        : 'Los sobrecupos son solo sábado y domingo. '+
          'Para otro día marca "Solicitud médica" y '+
          'adjunta el ticket (frente y reverso sellado '+
          'y firmado por el médico).'
    );
  }

  const frente=ticket3_(archivos && archivos.frente);
  const reverso=ticket3_(archivos && archivos.reverso);

  // El ticket es obligatorio en toda solicitud.
  if (!frente || !reverso) {
    fallo3_(
      'TICKET_REQUERIDO',
      'Adjunta el ticket de la consulta: frente y '+
      'reverso sellado y firmado por el médico.'
    );
  }

  if (finDeSemana &&
      limite3_(fecha,especialidad)===null) {
    fallo3_(
      'ESPECIALIDAD_NO_HABILITADA',
      'Esta especialidad no tiene sobrecupos '+
      'habilitados para ese día.'
    );
  }

  const id='UGM-'+Utilities.formatDate(
    new Date(),UGM3.TZ,'yyyyMMdd'
  )+'-'+Utilities.getUuid()
    .slice(0,8).toUpperCase();

  // Los archivos se guardan antes del bloqueo para no
  // detener a los demás; si algo falla se descartan.
  const guardados=[];

  const guardar=(t,etiqueta)=>{
    if (!t) return '';
    const archivo=carpeta3_().createFile(
      Utilities.newBlob(
        t.bytes,t.mime,
        id+'_'+etiqueta+'.'+EXTENSIONES3[t.mime]
      )
    );
    guardados.push(archivo);
    return archivo.getId();
  };

  let disponibles=null;
  let limiteFinal=null;
  let esUrgencia=false;
  const motivoUrgencia=texto3_(datos.motivoUrgencia)
    .replace(/\s+/g,' ').slice(0,300);

  try {
    const idFrente=guardar(frente,'ticket_frente');
    const idReverso=guardar(reverso,'ticket_reverso');

    const lock=LockService.getScriptLock();
    lock.waitLock(30000);

    try {
      const existentes=leer3_('SOBRECUPOS');

      const duplicado=existentes.find(r =>
        texto3_(r.FechaSobrecupo)===fecha &&
        norm3_(r.Especialidad)===
          norm3_(especialidad) &&
        texto3_(r.Documento)===documento &&
        !['RECHAZADO','ANULADO'].includes(
          norm3_(r.Estado)
        )
      );

      if (duplicado) {
        fallo3_(
          'DUPLICADO',
          'Ya existe una solicitud vigente: '+
          duplicado.ID
        );
      }

      if (finDeSemana) {
        // El cupo se valida dentro del bloqueo: dos
        // técnicos no pueden tomar el último a la vez.
        const limite=limite3_(fecha,especialidad) || 0;
        const ocupados=ocupados3_(
          existentes,fecha,especialidad
        );

        const extra=limite>0 ? UGM3.EXTRA_URGENCIA : 0;

        if (ocupados>=limite+extra) {
          fallo3_(
            'SIN_CUPOS',
            'Ya no se pueden agendar más sobrecupos de '+
            especialidad+' para el '+fecha+
            ' por restricción médica. Cupo completo: '+
            ocupados+' de '+limite+
            (extra ? ' + '+extra+' de urgencia' : '')+'.'
          );
        }

        if (ocupados>=limite) {
          // Solo queda el cupo de urgencia.
          if (datos.urgente!==true) {
            fallo3_(
              'SOLO_URGENCIA',
              'Los '+limite+' cupos regulares de '+
              especialidad+' para el '+fecha+' están '+
              'ocupados. Solo queda el cupo de urgencia: '+
              'marca el caso como urgente y escribe el motivo.'
            );
          }

          if (motivoUrgencia.length<UGM3.MIN_MOTIVO_URGENCIA) {
            fallo3_(
              'MOTIVO_URGENCIA',
              'Escribe el motivo de la urgencia (mínimo '+
              UGM3.MIN_MOTIVO_URGENCIA+' caracteres) para que '+
              'supervisión valide el caso.'
            );
          }

          esUrgencia=true;
        }

        disponibles=Math.max(0,limite-ocupados-1);
        limiteFinal=limite;
      }

      insertar3_('SOBRECUPOS',{
        ID:id,
        FechaSobrecupo:fecha,
        Especialidad:especialidad,
        Nombres:texto3_(datos.nombres).slice(0,120),
        Documento:documento,
        SolicitadoPor:
          texto3_(datos.solicitadoPor).slice(0,100),
        RegistradoPor:s.nombre,
        Sesion:s.id,
        FechaRegistro:ahora3_(),
        Estado:'PENDIENTE',
        ComprobanteID:idFrente,
        Tipo:medico ? 'MEDICO' : 'FIN_DE_SEMANA',
        TicketFrente:idFrente,
        TicketReverso:idReverso,
        Urgencia:esUrgencia ? 'SI' : '',
        MotivoUrgencia:esUrgencia ? motivoUrgencia : ''
      });

      SpreadsheetApp.flush();

    } finally {
      lock.releaseLock();
    }

  } catch(error) {
    guardados.forEach(a => {
      try { a.setTrashed(true); } catch(e) {}
    });
    throw error;
  }

  try {
    auditar3_(
      s.nombre,s.perfil,
      medico ? 'SOLICITUD_MEDICA' :
        esUrgencia ? 'CUPO_URGENCIA' : 'SOBRECUPO_GUARDADO',
      id,especialidad+' / '+fecha+
        (esUrgencia ? ' · URGENCIA: '+motivoUrgencia : '')
    );
  } catch(e) {
    console.error(e.message);
  }

  return {
    ok:true,
    id,
    disponibles,
    limite:limiteFinal,
    medico,
    urgencia:esUrgencia,
    mensaje:esUrgencia
      ? 'Cupo de urgencia registrado: supervisión debe validarlo.'
      : 'Solicitud registrada correctamente.'
  };
}

/* ==========================================================
   COMPROBANTES
   ========================================================== */

function bytesComprobante3_(mime,base64) {
  const permitidos=[
    'image/jpeg',
    'image/png',
    'image/webp',
    'application/pdf'
  ];

  if (!permitidos.includes(mime)) {
    throw Error('Formato no permitido.');
  }

  const bytes=Utilities.base64Decode(base64||'');

  if (!bytes.length ||
      bytes.length>UGM3.MAX_ARCHIVO_MB*1048576) {
    throw Error('Archivo vacío o mayor de 5 MB.');
  }

  const prefijo = a =>
    a.every((n,i)=>(bytes[i]&255)===n);

  const pdf=mime==='application/pdf' &&
    prefijo([37,80,68,70]);

  const png=mime==='image/png' &&
    prefijo([137,80,78,71]);

  const jpg=mime==='image/jpeg' &&
    prefijo([255,216]);

  const webp=mime==='image/webp' &&
    String.fromCharCode(
      ...bytes.slice(0,4)
    )==='RIFF' &&
    String.fromCharCode(
      ...bytes.slice(8,12)
    )==='WEBP';

  if (!(pdf||png||jpg||webp)) {
    throw Error(
      'El contenido del archivo no corresponde '+
      'al formato indicado.'
    );
  }

  return bytes;
}

function carpeta3_() {
  const p=props3_();
  const id=p.getProperty('UGM2_FOLDER_ID');

  if (id) return DriveApp.getFolderById(id);

  const carpeta=DriveApp.createFolder(
    'UGM 2026 - Comprobantes privados'
  );

  p.setProperty(
    'UGM2_FOLDER_ID',
    carpeta.getId()
  );

  return carpeta;
}

function adjuntarComprobante(
  token,id,nombre,mime,base64
) {
  const s=validar3_(token);

  if (!['SUPERVISOR','AUTORIZADO']
    .includes(s.perfil)) {
    throw Error(
      'No tienes permiso para adjuntar comprobantes.'
    );
  }

  const bytes=bytesComprobante3_(mime,base64);

  const lock=LockService.getScriptLock();
  lock.waitLock(30000);

  let archivo;

  try {
    const sh=hoja3_('SOBRECUPOS');
    const m=sh.getDataRange().getValues();

    const i=m.findIndex((r,pos)=>
      pos>0 && String(r[0])===String(id)
    );

    if (i<1) throw Error('Solicitud inexistente.');

    if (['REGISTRADO','RECHAZADO','ANULADO']
      .includes(texto3_(m[i][9]))) {
      throw Error(
        'No se puede adjuntar a una solicitud cerrada.'
      );
    }

    const nombreSeguro=texto3_(nombre)
      .replace(/[^a-zA-Z0-9_.-]/g,'_')
      .slice(0,80);

    archivo=carpeta3_().createFile(
      Utilities.newBlob(
        bytes,mime,id+'_'+nombreSeguro
      )
    );

    sh.getRange(i+1,13).setValue(
      archivo.getId()
    );

    SpreadsheetApp.flush();

  } finally {
    lock.releaseLock();
  }

  try {
    auditar3_(
      s.nombre,s.perfil,
      'COMPROBANTE_ADJUNTADO',
      id,archivo.getId()
    );
  } catch(e) {
    console.error(e.message);
  }

  return {
    ok:true,
    id:archivo.getId(),
    nombre:archivo.getName()
  };
}

/* ==========================================================
   CONFIRMACIÓN / CAMBIO DE ESTADO
   ========================================================== */

function cambiarEstado(
  token,id,nuevoEstado,observacion
) {
  const s=validar3_(token);

  if (!['SUPERVISOR','AUTORIZADO']
    .includes(s.perfil)) {
    throw Error(
      'No tienes permisos de confirmación.'
    );
  }

  const permitido=[
    'EN GESTION',
    'OBSERVADO',
    'VALIDACION',
    'REGISTRADO',
    'RECHAZADO',
    'ANULADO'
  ];

  if (!permitido.includes(nuevoEstado)) {
    throw Error('Estado inválido.');
  }

  const lock=LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sh=hoja3_('SOBRECUPOS');
    const m=sh.getDataRange().getValues();

    const i=m.findIndex((r,pos)=>
      pos>0 && String(r[0])===String(id)
    );

    if (i<1) {
      throw Error('Solicitud inexistente.');
    }

    const actual=texto3_(m[i][9]);

    // Una solicitud agendada solo puede anularse para
    // liberar su cupo, indicando el motivo.
    const liberar=actual==='REGISTRADO' &&
      nuevoEstado==='ANULADO';

    if (liberar && !texto3_(observacion)) {
      throw Error('Indica el motivo para liberar el cupo.');
    }

    if (!liberar && ['REGISTRADO','RECHAZADO','ANULADO']
      .includes(actual)) {
      throw Error('La solicitud ya está cerrada.');
    }

    if (nuevoEstado==='REGISTRADO') {
      const comprobante=texto3_(m[i][12]) ||
        texto3_(m[i][columna3_('SOBRECUPOS','TicketFrente')]);

      if (!comprobante) {
        throw Error(
          'Es obligatorio adjuntar un comprobante '+
          'antes de confirmar.'
        );
      }

      try {
        const archivo=DriveApp
          .getFileById(comprobante);

        if (archivo.isTrashed()) {
          throw Error('Archivo no disponible.');
        }
      } catch(e) {
        throw Error(
          'No se pudo verificar el comprobante.'
        );
      }
    }

    sh.getRange(i+1,10,1,5)
      .setValues([[
        nuevoEstado,
        s.nombre,
        ahora3_(),
        m[i][12]||'',
        seguro3_(
          texto3_(observacion).slice(0,500)
        )
      ]]);

    SpreadsheetApp.flush();

  } finally {
    lock.releaseLock();
  }

  try {
    auditar3_(
      s.nombre,s.perfil,
      'CAMBIO_ESTADO',
      id,nuevoEstado
    );
  } catch(e) {
    console.error(e.message);
  }

  return {ok:true,estado:nuevoEstado};
}

/* ==========================================================
   BASE DE PACIENTES
   La información de la base se guarda en la hoja
   oculta UGM_CONFIG, así la ven el supervisor y los
   técnicos aunque usen implementaciones distintas
   del mismo libro.
   ========================================================== */

const HOJA_CONFIG3='UGM_CONFIG';

function config3_() {
  const sh=base3_().getSheetByName(HOJA_CONFIG3);
  const datos={};

  if (!sh || sh.getLastRow()<2) return datos;

  sh.getRange(2,1,sh.getLastRow()-1,2)
    .getValues()
    .forEach(r => {
      if (r[0]!=='') datos[String(r[0])]=String(r[1]);
    });

  return datos;
}

function guardarConfig3_(valores) {
  const libro=base3_();
  let sh=libro.getSheetByName(HOJA_CONFIG3);

  if (!sh) {
    sh=libro.insertSheet(HOJA_CONFIG3);
    sh.getRange(1,1,1,2).setValues([['Clave','Valor']]);
    sh.getRange('B:B').setNumberFormat('@');
    sh.hideSheet();
  }

  const ultima=sh.getLastRow();
  const claves=ultima>1
    ? sh.getRange(2,1,ultima-1,1).getValues()
        .map(r=>String(r[0]))
    : [];

  Object.keys(valores).forEach(k => {
    let i=claves.indexOf(k);

    if (i<0) {
      claves.push(k);
      i=claves.length-1;
    }

    sh.getRange(i+2,1,1,2)
      .setValues([[k,String(valores[k])]]);
  });
}

function hojaPacientes3_(libro,nombre) {
  const sh=libro.insertSheet(nombre);
  const columnas=UGM3.HEADERS.PACIENTES.length;

  sh.getRange(1,1,1,columnas).setValues([
    UGM3.HEADERS.PACIENTES
  ]);

  sh.getRange(1,1,1,columnas)
    .setBackground('#10456B')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  sh.setFrozenRows(1);
  sh.getRange('B:E').setNumberFormat('@');
  sh.getRange('G:H').setNumberFormat('@');

  return sh;
}

// Ubica las columnas por encabezado.
function columnasPacientes3_(sh) {
  const cab=sh.getRange(1,1,1,sh.getLastColumn())
    .getValues()[0].map(String);

  const i=n=>cab.indexOf(n);

  return {
    nombre:i('Nombre'),
    documento:i('Documento'),
    original:i('DocumentoOriginal'),
    expediente:i('Expediente'),
    especialidad:i('Especialidad'),
    busqueda:i('Busqueda'),
    fecha:i('Fecha'),
    hora:i('Hora'),
    aumento:i('Aumento'),
    observacion:i('ObservacionUCC'),
    correlativo:i('Correlativo'),
    ancho:cab.length
  };
}

function estadoPacientes3_() {
  const c=config3_();

  return {
    version:c.PAC_VERSION || '',
    fecha:c.PAC_FECHA || '',
    archivo:c.PAC_ARCHIVO || '',
    usuario:c.PAC_USUARIO || '',
    total:Number(c.PAC_TOTAL) || 0,
    invalidos:Number(c.PAC_INVALIDOS) || 0,
    conEspecialidad:c.PAC_ESPECIALIDAD==='1',
    conFechas:c.PAC_CONFECHAS==='1',
    // {"2026-10-10":{"total":387,"aumento":42}, ...}
    fechas:(()=>{
      try { return JSON.parse(c.PAC_FECHAS || '{}'); }
      catch(e) { return {}; }
    })()
  };
}

function especialidadesTexto3_(valor) {
  return texto3_(valor)
    .split(/[;,|]/)
    .map(e=>texto3_(e))
    .filter(Boolean)
    .map(e=>especialidad3_(e) || e.slice(0,80))
    .filter((e,i,a)=>a.indexOf(e)===i)
    .join(' | ')
    .slice(0,300);
}

/**
 * Carga la base de pacientes por lotes (solo
 * supervisor). El primer lote crea una hoja temporal;
 * el último la publica como PACIENTES, de modo que una
 * carga interrumpida nunca deja la base a medias.
 */
function cargarPacientes(
  token,carga,lote,parte,partes,archivo
) {
  const s=validar3_(token);

  if (s.perfil!=='SUPERVISOR') {
    throw Error(
      'Solo el supervisor puede cargar la base de pacientes.'
    );
  }

  parte=Number(parte);
  partes=Number(partes);

  if (!Number.isInteger(parte) ||
      !Number.isInteger(partes) ||
      parte<0 || parte>=partes ||
      partes*UGM3.LOTE_PACIENTES >
        UGM3.MAX_PACIENTES+UGM3.LOTE_PACIENTES) {
    throw Error('Lote de carga inválido.');
  }

  if (!Array.isArray(lote) ||
      lote.length>UGM3.LOTE_PACIENTES) {
    throw Error('Lote de pacientes inválido.');
  }

  let invalidos=0;
  let conEspecialidad=false;
  const fechasLote={};

  const filas=lote.map(r => {
    r=r || {};

    const nombre=texto3_(r.nombre)
      .replace(/\s+/g,' ')
      .slice(0,150);

    if (!nombre) return null;

    const original=texto3_(r.documento).slice(0,30);
    const documento=dui3_(original);
    const especialidad=especialidadesTexto3_(
      r.especialidad
    );

    if (!documento) invalidos++;
    if (especialidad) conEspecialidad=true;

    // Fecha de la cita (Reportería): aaaa-mm-dd.
    const fecha=/^\d{4}-\d{2}-\d{2}$/.test(texto3_(r.fecha))
      ? texto3_(r.fecha) : '';
    const aumento=r.aumento===true || texto3_(r.aumento)==='SI';

    if (fecha) {
      const f=fechasLote[fecha]=fechasLote[fecha] ||
        {total:0,aumento:0};
      f.total++;
      if (aumento) f.aumento++;
    }

    return [
      seguro3_(nombre),
      documento,
      seguro3_(original),
      seguro3_(texto3_(r.expediente).slice(0,50)),
      seguro3_(especialidad),
      busqueda3_(nombre),
      fecha,
      texto3_(r.hora).slice(0,10),
      aumento ? 'SI' : '',
      seguro3_(texto3_(r.observacion).replace(/\s+/g,' ').slice(0,200)),
      seguro3_(texto3_(r.correlativo).slice(0,40))
    ];
  }).filter(Boolean);

  const lock=LockService.getScriptLock();
  lock.waitLock(30000);

  let final=null;

  try {
    const p=props3_();
    const libro=base3_();
    const TEMPORAL='PACIENTES_CARGA';

    if (parte===0) {
      carga=Utilities.getUuid();

      const previa=libro.getSheetByName(TEMPORAL);
      if (previa) libro.deleteSheet(previa);

      hojaPacientes3_(libro,TEMPORAL);

      p.setProperties({
        UGM3_PAC_CARGA:carga,
        UGM3_PAC_CARGA_INV:'0',
        UGM3_PAC_CARGA_ESP:'',
        UGM3_PAC_CARGA_FEC:'{}'
      });

    } else if (
      !carga ||
      p.getProperty('UGM3_PAC_CARGA')!==String(carga)
    ) {
      throw Error(
        'La carga de pacientes se interrumpió o fue '+
        'reemplazada por otra. Vuelve a intentarlo.'
      );
    }

    const sh=libro.getSheetByName(TEMPORAL);

    if (!sh) {
      throw Error('No existe la hoja temporal de carga.');
    }

    if (filas.length) {
      sh.getRange(
        sh.getLastRow()+1,1,
        filas.length,filas[0].length
      ).setValues(filas);
    }

    const acumulados=
      Number(p.getProperty('UGM3_PAC_CARGA_INV')||0)+
      invalidos;

    const tieneEspecialidad=conEspecialidad ||
      p.getProperty('UGM3_PAC_CARGA_ESP')==='1';

    let fechas={};
    try {
      fechas=JSON.parse(p.getProperty('UGM3_PAC_CARGA_FEC') || '{}');
    } catch(e) {}

    Object.keys(fechasLote).forEach(f => {
      const a=fechas[f]=fechas[f] || {total:0,aumento:0};
      a.total+=fechasLote[f].total;
      a.aumento+=fechasLote[f].aumento;
    });

    p.setProperties({
      UGM3_PAC_CARGA_INV:String(acumulados),
      UGM3_PAC_CARGA_ESP:tieneEspecialidad ? '1' : '',
      UGM3_PAC_CARGA_FEC:JSON.stringify(fechas)
    });

    if (parte===partes-1) {
      SpreadsheetApp.flush();

      const total=sh.getLastRow()-1;

      if (total<1) {
        libro.deleteSheet(sh);
        throw Error(
          'No se encontraron pacientes con nombre.'
        );
      }

      const anterior=libro.getSheetByName('PACIENTES');
      if (anterior) libro.deleteSheet(anterior);

      sh.setName('PACIENTES');

      guardarConfig3_({
        PAC_VERSION:carga,
        PAC_FECHA:ahora3_(),
        PAC_ARCHIVO:texto3_(archivo).slice(0,150),
        PAC_USUARIO:s.nombre,
        PAC_TOTAL:total,
        PAC_INVALIDOS:acumulados,
        PAC_ESPECIALIDAD:tieneEspecialidad ? '1' : '0',
        PAC_CONFECHAS:Object.keys(fechas).length ? '1' : '0',
        PAC_FECHAS:JSON.stringify(fechas)
      });

      p.setProperty('UGM3_PAC_CARGA','');

      final={total,invalidos:acumulados,fechas};
    }

  } finally {
    lock.releaseLock();
  }

  if (final) {
    try {
      auditar3_(
        s.nombre,s.perfil,
        'PACIENTES_IMPORTADOS','',
        texto3_(archivo)+' · '+final.total+
        ' pacientes · DUI inválidos: '+final.invalidos
      );
    } catch(e) {
      console.error(e.message);
    }
  }

  return {
    ok:true,
    carga,
    final:Boolean(final),
    total:final ? final.total : 0,
    invalidos:final ? final.invalidos : 0,
    fechas:final ? final.fechas : {}
  };
}

/**
 * Entrega la base por páginas para que el navegador
 * busque al instante, sin consultar al servidor en
 * cada letra.
 */
function descargarPacientes(token,version,desde) {
  validar3_(token);

  const estado=estadoPacientes3_();

  if (!estado.version) {
    return {version:'',total:0,filas:[],fin:true};
  }

  if (version && version!==estado.version) {
    return {cambio:true,version:estado.version};
  }

  const sh=base3_().getSheetByName('PACIENTES');

  if (!sh || sh.getLastRow()<2) {
    return {
      version:estado.version,total:0,filas:[],fin:true
    };
  }

  const total=sh.getLastRow()-1;
  desde=Math.max(0,Math.floor(Number(desde)||0));

  const n=Math.min(UGM3.PAGINA_PACIENTES,total-desde);

  if (n<=0) {
    return {
      version:estado.version,total,filas:[],fin:true
    };
  }

  const c=columnasPacientes3_(sh);
  const v=(r,i)=>i>=0 ? texto3_(r[i]) : '';

  const filas=sh.getRange(2+desde,1,n,c.ancho)
    .getValues()
    .map(r=>[
      v(r,c.nombre),
      v(r,c.documento),
      v(r,c.original),
      v(r,c.expediente),
      v(r,c.especialidad),
      v(r,c.fecha),
      v(r,c.hora),
      v(r,c.aumento)
    ]);

  return {
    version:estado.version,
    total,
    desde,
    filas,
    fin:desde+n>=total
  };
}

/**
 * Especialidades de la agenda por paciente, para
 * buscar pacientes por especialidad aunque la base
 * no traiga esa columna.
 */
function mapaAgendaPacientes(token) {
  validar3_(token);

  const vistos={};
  const filas=[];

  leer3_('AGENDA_EXCEL').forEach(r => {
    const esp=texto3_(r.Especialidad);
    if (!esp || excluida3_(esp)) return;

    const doc=texto3_(r.Documento);
    const dui=dui3_(doc);
    const alfa=norm3_(doc).replace(/[^A-Z0-9]/g,'');
    const nombre=busqueda3_(r.Paciente);

    if (!dui && !alfa && !nombre) return;

    const clave=[dui,alfa,nombre,esp].join('|');
    if (vistos[clave]) return;
    vistos[clave]=true;

    filas.push([dui,alfa,nombre,esp]);
  });

  return {
    version:estadoAgenda3_().cargaID,
    filas
  };
}

/**
 * Búsqueda en el servidor: solo se usa mientras el
 * navegador termina de descargar la base.
 */
function buscarPacientes(token,consulta) {
  validar3_(token);

  const texto=texto3_(consulta).slice(0,80);
  const q=busqueda3_(texto);
  const digitos=texto.replace(/\D/g,'');

  const porDui=digitos.length>=4 &&
    /^[0-9\s-]+$/.test(texto);

  if (!porDui && q.replace(/ /g,'').length<3) {
    return {resultados:[]};
  }

  const sh=base3_().getSheetByName('PACIENTES');

  if (!sh || sh.getLastRow()<2) {
    return {resultados:[],sinBase:true};
  }

  const c=columnasPacientes3_(sh);
  const v=(r,i)=>i>=0 ? String(r[i]) : '';

  const matriz=sh.getRange(
    2,1,sh.getLastRow()-1,c.ancho
  ).getValues();

  const palabras=q.split(' ').filter(Boolean);
  const inicio=palabras[0] || '';
  const encontrados=[];

  for (let i=0;i<matriz.length;i++) {
    const r=matriz[i];
    const b=c.busqueda>=0
      ? v(r,c.busqueda) : busqueda3_(v(r,c.nombre));
    let coincide;

    if (porDui) {
      coincide=
        v(r,c.documento).replace(/\D/g,'')
          .startsWith(digitos) ||
        v(r,c.original).replace(/\D/g,'')
          .startsWith(digitos);
    } else {
      coincide=palabras.every(t=>b.includes(t));
    }

    if (!coincide) continue;

    const original=v(r,c.original);
    const documento=v(r,c.documento);
    const largo=original
      .replace(/\.0+$/,'')
      .replace(/\D/g,'').length;

    encontrados.push({
      nombre:v(r,c.nombre),
      documento,
      valido:Boolean(documento),
      ajustado:Boolean(documento) && largo===8,
      original,
      expediente:v(r,c.expediente),
      especialidades:v(r,c.especialidad)
        .split(' | ').filter(Boolean),
      cita:c.fecha>=0 && v(r,c.fecha) ? {
        fecha:v(r,c.fecha),
        hora:v(r,c.hora),
        aumento:v(r,c.aumento)==='SI'
      } : null,
      prioridad:b.startsWith(inicio) ? 0 : 1
    });

    if (encontrados.length>=200) break;
  }

  encontrados.sort((a,b)=>a.prioridad-b.prioridad);

  return {
    resultados:encontrados
      .slice(0,UGM3.MAX_RESULTADOS)
      .map(({prioridad,...r})=>r)
  };
}

/* ==========================================================
   CONSULTAS DEL DÍA Y ASISTENCIA
   ========================================================== */

function claveAsistencia3_(fecha,especialidad,documento,paciente) {
  const doc=norm3_(documento).replace(/[^A-Z0-9]/g,'');
  const id=doc ? 'D:'+doc : 'N:'+busqueda3_(paciente);
  return fecha+'|'+norm3_(especialidad)+'|'+id;
}

/**
 * Citas de una fecha (agenda + sobrecupos vigentes),
 * una por paciente y especialidad, con su asistencia.
 */
/**
 * Pacientes de la base de Reportería con cita en la
 * fecha (solo si la base trae la columna de fecha). Se
 * guarda en caché unos minutos por versión de la base,
 * porque las pantallas se sincronizan cada 10 s.
 */
function pacientesDelDia3_(fecha) {
  const estado=estadoPacientes3_();
  if (!estado.conFechas || !estado.fechas[fecha]) return [];

  const clave='PACDIA_'+estado.version.slice(0,12)+'_'+fecha;
  const cache=CacheService.getScriptCache();

  try {
    const guardado=cache.get(clave);
    if (guardado) return JSON.parse(guardado);
  } catch(e) {}

  const sh=base3_().getSheetByName('PACIENTES');
  if (!sh || sh.getLastRow()<2) return [];

  const c=columnasPacientes3_(sh);
  if (c.fecha<0) return [];

  const v=(r,i)=>i>=0 ? texto3_(r[i]) : '';

  const lista=sh.getRange(2,1,sh.getLastRow()-1,c.ancho)
    .getValues()
    .filter(r=>v(r,c.fecha)===fecha)
    .map(r=>({
      nombre:v(r,c.nombre),
      documento:v(r,c.documento),
      original:v(r,c.original),
      especialidad:v(r,c.especialidad).split(' | ')[0] || '',
      hora:v(r,c.hora),
      aumento:v(r,c.aumento)==='SI',
      observacion:v(r,c.observacion),
      expediente:v(r,c.expediente).replace(/\.0+$/,''),
      correlativo:v(r,c.correlativo)
    }));

  try {
    const json=JSON.stringify(lista);
    if (json.length<95000) cache.put(clave,json,300);
  } catch(e) {}

  return lista;
}

// Regla de jefatura: los domingos solo hay pacientes
// extraordinarios (no se programa consulta ordinaria).
function esDomingo3_(fecha) {
  const d=new Date(texto3_(fecha)+'T12:00:00Z');
  return !isNaN(d.getTime()) && d.getUTCDay()===0;
}

function esSabado3_(fecha) {
  const d=new Date(texto3_(fecha)+'T12:00:00Z');
  return !isNaN(d.getTime()) && d.getUTCDay()===6;
}

// Clave para comparar especialidades sin tildes, espacios ni
// guiones ("Medicina Familiar-Diabetología" =
// "Medicina Familiar - Diabetología"). La comparación es exacta:
// "Cardiología - Estudios Especiales" NO es "Cardiología".
function claveEsp3_(v) {
  return norm3_(v).replace(/[^A-Z0-9]/g,'');
}

// Especialidades extraordinarias del sábado (lista del
// programa de fin de semana; se actualiza aquí cada semana).
function listaSabado3_() {
  return UGM3.ESPECIALIDADES_SABADO;
}

// Extraordinario:
// - DOMINGO: todas las citas.
// - SÁBADO: la especialidad está en la lista del sábado.
// Todo lo demás es ordinario. Los sobrecupos de técnicos se
// muestran aparte.
function clasificar3_(c, fecha) {
  let motivo='';
  if (c.origen!=='SOBRECUPO') {
    if (esDomingo3_(fecha)) {
      motivo='Cita en domingo';
    } else if (esSabado3_(fecha)) {
      const set=new Set(listaSabado3_().map(claveEsp3_));
      if (set.has(claveEsp3_(c.especialidad))) {
        motivo='Especialidad del programa del sábado';
      }
    }
  }
  c.extraordinario=Boolean(motivo);
  c.motivoExtra=motivo;
  return c;
}

function consultasDelDia3_(fecha,agenda,solicitudes,asistencia,reporteria) {
  const mapa=new Map();

  // Citas canceladas en la agenda del sistema: si el
  // paciente viene en la base de Reportería, no se muestra.
  const canceladas={codigos:{},pacientes:{}};

  agenda.forEach(r => {
    if (texto3_(r.Fecha)!==fecha) return;
    if (excluida3_(r.Especialidad)) return;

    const estado=estadoCita3_(r.Estado);
    if (estado==='CANCELADA') {
      if (texto3_(r.IDExterno)) canceladas.codigos[texto3_(r.IDExterno)]=true;
      const d=dui3_(r.Documento);
      if (d) canceladas.pacientes[d+'|'+norm3_(r.Especialidad)]=true;
      canceladas.pacientes['N:'+busqueda3_(r.Paciente)+'|'+
        norm3_(r.Especialidad)]=true;
      return;
    }

    const clave=claveAsistencia3_(
      fecha,r.Especialidad,r.Documento,r.Paciente
    );

    if (mapa.has(clave)) {
      if (estado==='ACTIVA') {
        mapa.get(clave).estadoCita='ACTIVA';
      }
      return;
    }

    mapa.set(clave,{
      clave,
      especialidad:texto3_(r.Especialidad),
      paciente:texto3_(r.Paciente),
      documento:texto3_(r.Documento),
      estadoCita:estado,
      origen:'AGENDA',
      referencia:texto3_(r.IDExterno),
      creacion:texto3_(r.FechaCreacion),
      expediente:texto3_(r.Expediente),
      sobrecupo:''
    });
  });

  // Pacientes de Reportería (ya agendados para el fin de
  // semana). Si ya están en la agenda del sistema no se
  // duplican: solo se les agrega la hora y la marca de
  // aumento del 10 %.
  const porDui={};
  const porNombre={};
  const porCodigo={};

  mapa.forEach(c => {
    if (c.referencia) porCodigo[c.referencia]=c;
    const d=dui3_(c.documento);
    if (d) porDui[d+'|'+norm3_(c.especialidad)]=c;
    porNombre[busqueda3_(c.paciente)+'|'+norm3_(c.especialidad)]=c;
  });

  (reporteria || []).forEach(p => {
    if (excluida3_(p.especialidad)) return;
    const esp=norm3_(p.especialidad);
    const dui=p.documento || dui3_(p.original);
    const cod=texto3_(p.correlativo);

    // 1) mismo número de cita (CORRELATIVO = Codigo de la
    //    agenda); 2) mismo DUI; 3) mismo nombre. Siempre en
    //    la misma especialidad, salvo el número de cita.
    const existente=(cod && porCodigo[cod]) ||
      (dui && porDui[dui+'|'+esp]) ||
      porNombre[busqueda3_(p.nombre)+'|'+esp];

    if (!existente && (
        (cod && canceladas.codigos[cod]) ||
        (dui && canceladas.pacientes[dui+'|'+esp]) ||
        canceladas.pacientes['N:'+busqueda3_(p.nombre)+'|'+esp])) {
      return;
    }

    if (existente) {
      existente.reporteria=true;
      existente.aumento=existente.aumento || p.aumento;
      existente.horaCita=existente.horaCita || p.hora;
      existente.observacion=existente.observacion || p.observacion;
      existente.expediente=existente.expediente || p.expediente;
      return;
    }

    const documento=dui || texto3_(p.original);
    const clave=claveAsistencia3_(
      fecha,p.especialidad,documento,p.nombre
    );

    if (mapa.has(clave)) return;

    const nueva={
      clave,
      especialidad:texto3_(p.especialidad),
      paciente:texto3_(p.nombre),
      documento,
      estadoCita:'ACTIVA',
      origen:'REPORTERIA',
      referencia:'',
      sobrecupo:'',
      reporteria:true,
      aumento:p.aumento,
      horaCita:p.hora,
      observacion:p.observacion,
      expediente:p.expediente
    };

    mapa.set(clave,nueva);
    if (dui) porDui[dui+'|'+esp]=nueva;
    porNombre[busqueda3_(p.nombre)+'|'+esp]=nueva;
  });

  solicitudes.forEach(r => {
    if (texto3_(r.FechaSobrecupo)!==fecha) return;

    if (['RECHAZADO','ANULADO'].includes(
      norm3_(r.Estado)
    )) return;

    const clave=claveAsistencia3_(
      fecha,r.Especialidad,r.Documento,r.Nombres
    );

    // El mismo paciente puede venir con el DUI escrito
    // distinto en la agenda: se busca también por DUI.
    const igual=mapa.get(clave) ||
      porDui[dui3_(r.Documento)+'|'+norm3_(r.Especialidad)];

    if (igual) {
      igual.sobrecupo=texto3_(r.Estado);
      return;
    }

    mapa.set(clave,{
      clave,
      especialidad:texto3_(r.Especialidad),
      paciente:texto3_(r.Nombres),
      documento:texto3_(r.Documento),
      estadoCita:'SOBRECUPO',
      origen:'SOBRECUPO',
      referencia:texto3_(r.ID),
      sobrecupo:texto3_(r.Estado)
    });
  });

  const marcas={};

  asistencia.forEach(a => {
    if (excluida3_(a.Especialidad)) return;
    if (texto3_(a.Fecha)===fecha) {
      marcas[texto3_(a.Clave)]=a;
    }
  });

  // Marcas cuya cita ya no está en la agenda cargada:
  // se conservan para el reporte.
  Object.keys(marcas).forEach(clave => {
    if (mapa.has(clave)) return;

    const a=marcas[clave];

    mapa.set(clave,{
      clave,
      especialidad:texto3_(a.Especialidad),
      paciente:texto3_(a.Paciente),
      documento:texto3_(a.Documento),
      estadoCita:'REGISTRADA',
      origen:texto3_(a.Origen) || 'REGISTRO',
      referencia:texto3_(a.Referencia),
      sobrecupo:''
    });
  });

  return [...mapa.values()]
    .map(c => {
      clasificar3_(c, fecha);
      const a=marcas[c.clave];
      c.asistencia=a ? texto3_(a.Asistencia) : '';
      c.marcadoPor=a ? texto3_(a.MarcadoPor) : '';
      c.hora=a ? texto3_(a.FechaHora) : '';
      return c;
    })
    .sort((a,b) =>
      a.especialidad.localeCompare(b.especialidad) ||
      String(a.horaCita||'99').localeCompare(String(b.horaCita||'99')) ||
      a.paciente.localeCompare(b.paciente)
    );
}

function resumenAsistencia3_(consultas) {
  const porEsp={};
  const total={citas:0,asistio:0,noAsistio:0,sinMarcar:0};

  consultas.forEach(c => {
    const e=porEsp[c.especialidad] ||
      (porEsp[c.especialidad]={
        especialidad:c.especialidad,
        citas:0,asistio:0,noAsistio:0,sinMarcar:0
      });

    const campo=c.asistencia==='ASISTIO' ? 'asistio' :
      c.asistencia==='NO_ASISTIO' ? 'noAsistio' :
      'sinMarcar';

    e.citas++; e[campo]++;
    total.citas++; total[campo]++;
  });

  return {
    total,
    especialidades:Object.values(porEsp)
      .sort((a,b)=>b.citas-a.citas)
  };
}

function marcarAsistencia(token,datos) {
  const s=validar3_(token);

  datos=datos || {};

  const fecha=fecha3_(datos.fecha);
  const especialidad=texto3_(datos.especialidad).slice(0,100);
  const paciente=texto3_(datos.paciente).slice(0,150);
  const documento=texto3_(datos.documento).slice(0,40);
  const valor=texto3_(datos.asistencia).toUpperCase();

  if (!['ASISTIO','NO_ASISTIO',''].includes(valor)) {
    throw Error('Valor de asistencia inválido.');
  }

  if (!especialidad || (!paciente && !documento)) {
    throw Error('Datos de la cita incompletos.');
  }

  if (fecha>hoy3_()) {
    throw Error(
      'La asistencia se marca el día de la cita.'
    );
  }

  const clave=claveAsistencia3_(
    fecha,especialidad,documento,paciente
  );

  const hora=ahora3_();

  const fila=[
    clave,fecha,especialidad,documento,paciente,
    texto3_(datos.origen).slice(0,20),
    texto3_(datos.referencia).slice(0,100),
    valor,s.nombre,s.id,hora
  ].map(seguro3_);

  const lock=LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const sh=hoja3_('ASISTENCIA');
    const ultima=sh.getLastRow();

    const claves=ultima>1
      ? sh.getRange(2,1,ultima-1,1).getValues()
      : [];

    const i=claves.findIndex(
      r=>String(r[0])===clave
    );

    sh.getRange(
      i>=0 ? i+2 : ultima+1,1,1,fila.length
    ).setValues([fila]);

    SpreadsheetApp.flush();

  } finally {
    lock.releaseLock();
  }

  try {
    auditar3_(
      s.nombre,s.perfil,
      valor ? 'ASISTENCIA_'+valor : 'ASISTENCIA_BORRADA',
      paciente || documento,
      especialidad+' / '+fecha
    );
  } catch(e) {
    console.error(e.message);
  }

  return {
    ok:true,
    clave,
    asistencia:valor,
    marcadoPor:s.nombre,
    hora
  };
}

/* ==========================================================
   PLATAFORMA DE TÉCNICOS
   ========================================================== */

function tableroTecnico(token,fecha,actividad) {
  const s=validar3_(token);
  const f=fecha3_(fecha);

  actividad3_(s,actividad);

  const agenda=leer3_('AGENDA_EXCEL');
  const solicitudes=leer3_('SOBRECUPOS');
  const asistencia=leer3_('ASISTENCIA');

  const consultas=consultasDelDia3_(
    f,agenda,solicitudes,asistencia,pacientesDelDia3_(f)
  );

  const nombre=norm3_(s.nombre);

  const mias=solicitudes
    .filter(r =>
      String(r.Sesion)===s.id ||
      norm3_(r.RegistradoPor)===nombre
    )
    .sort((a,b)=>
      String(b.FechaRegistro||'').localeCompare(
        String(a.FechaRegistro||'')
      )
    )
    .slice(0,100)
    .map(r=>({
      ID:r.ID,
      FechaSobrecupo:r.FechaSobrecupo,
      Especialidad:r.Especialidad,
      Nombres:r.Nombres,
      Documento:r.Documento,
      Estado:r.Estado,
      Tipo:r.Tipo || 'FIN_DE_SEMANA',
      Observacion:r.Observacion,
      FechaRegistro:r.FechaRegistro,
      ticket:Boolean(r.TicketFrente || r.ComprobanteID)
    }));

  return {
    fecha:f,
    hoy:hoy3_(),
    dia:dia3_(f),
    bloqueada:bloqueada3_(f),
    agenda:estadoAgenda3_(),
    pacientes:estadoPacientes3_(),
    consultas,
    resumen:resumenAsistencia3_(consultas).total,
    solicitudes:mias
  };
}

function consultarCupos(token,fecha) {
  validar3_(token);

  const f=fecha3_(fecha);

  return {
    fecha:f,
    dia:dia3_(f),
    bloqueada:bloqueada3_(f),
    agenda:estadoAgenda3_(),
    lista:calcular3_(
      f,leer3_('AGENDA_EXCEL'),leer3_('SOBRECUPOS')
    )
  };
}

/* ==========================================================
   PLATAFORMA DE SUPERVISIÓN
   ========================================================== */

function supervisor3_(token) {
  const s=validar3_(token);

  if (!['SUPERVISOR','AUTORIZADO'].includes(s.perfil)) {
    throw Error('Acceso exclusivo de supervisión.');
  }

  return s;
}

function tableroSupervisor(token,fecha,opciones) {
  const s=supervisor3_(token);
  const f=fecha3_(fecha);
  opciones=opciones || {};

  actividad3_(s,opciones.actividad || 'Centro de control');

  const agenda=leer3_('AGENDA_EXCEL');
  const solicitudes=leer3_('SOBRECUPOS');
  const asistencia=leer3_('ASISTENCIA');
  const personal=leer3_('PERSONAL');
  const auditoria=leer3_('AUDITORIA');
  const hoy=hoy3_();
  const ahora=Date.now();

  const consultas=consultasDelDia3_(
    f,agenda,solicitudes,asistencia,pacientesDelDia3_(f)
  );

  const cupos=calcular3_(f,agenda,solicitudes);

  const estados={};
  const estadosFecha={};
  let registradasHoy=0;
  let medicasPendientes=0;

  solicitudes.forEach(r => {
    const e=texto3_(r.Estado) || 'PENDIENTE';
    estados[e]=(estados[e]||0)+1;

    if (texto3_(r.FechaSobrecupo)===f) {
      estadosFecha[e]=(estadosFecha[e]||0)+1;
    }

    if (String(r.FechaRegistro).slice(0,10)===hoy ||
        (r.FechaRegistro &&
         Utilities.formatDate(
           new Date(r.FechaRegistro),UGM3.TZ,'yyyy-MM-dd'
         )===hoy)) {
      registradasHoy++;
    }

    if (r.Tipo==='MEDICO' &&
        ['PENDIENTE','EN GESTION','OBSERVADO','VALIDACION']
          .includes(norm3_(r.Estado))) {
      medicasPendientes++;
    }
  });

  const sesiones=personal
    .filter(r=>r.Estado==='CONECTADO')
    .map(r => {
      const ultima=Date.parse(r.UltimaActividad);
      const segundos=isNaN(ultima)
        ? null : Math.round((ahora-ultima)/1000);
      return {
        Sesion:r.Sesion,
        Nombre:r.Nombre,
        Perfil:r.Perfil,
        Ingreso:r.Ingreso,
        UltimaActividad:r.UltimaActividad,
        Actividad:r.Actividad,
        segundos,
        enLinea:segundos!==null && segundos<90
      };
    })
    // Inactivas por más de 12 h ya vencieron.
    .filter(r=>r.segundos!==null && r.segundos<43200)
    .sort((a,b)=>a.segundos-b.segundos);

  const horas=new Array(24).fill(0);

  auditoria.forEach(r => {
    const d=new Date(r.FechaHora);
    if (isNaN(d.getTime())) return;
    if (Utilities.formatDate(d,UGM3.TZ,'yyyy-MM-dd')!==hoy) {
      return;
    }
    horas[Number(Utilities.formatDate(d,UGM3.TZ,'H'))]++;
  });

  const resumen=resumenAsistencia3_(consultas);

  const ordenadas=solicitudes.slice().sort((a,b)=>
    String(b.FechaRegistro||'').localeCompare(
      String(a.FechaRegistro||'')
    )
  );

  return {
    fecha:f,
    hoy,
    dia:dia3_(f),
    bloqueada:bloqueada3_(f),
    agenda:estadoAgenda3_(),
    pacientes:estadoPacientes3_(),
    cupos,
    estados,
    estadosFecha,
    registradasHoy,
    medicasPendientes,
    pendientes:solicitudes.filter(r =>
      ['PENDIENTE','EN GESTION','OBSERVADO','VALIDACION']
        .includes(norm3_(r.Estado))
    ).length,
    bandeja:bandeja3_(solicitudes),
    limitesGenerales:opciones.limites ? limitesGenerales3_() : [],
    personal:sesiones,
    auditoria:auditoria.slice(-80).reverse(),
    horas,
    asistencia:resumen,
    consultas:opciones.detalle ? consultas : [],
    solicitudes:ordenadas.slice(0,400).map(r=>({
      ID:r.ID,
      FechaSobrecupo:r.FechaSobrecupo,
      Especialidad:r.Especialidad,
      Nombres:r.Nombres,
      Documento:r.Documento,
      SolicitadoPor:r.SolicitadoPor,
      RegistradoPor:r.RegistradoPor,
      FechaRegistro:r.FechaRegistro,
      Estado:r.Estado,
      GestionadoPor:r.GestionadoPor,
      Observacion:r.Observacion,
      Tipo:r.Tipo || 'FIN_DE_SEMANA',
      ComprobanteID:r.ComprobanteID,
      TicketFrente:r.TicketFrente,
      TicketReverso:r.TicketReverso,
      Urgencia:texto3_(r.Urgencia),
      MotivoUrgencia:texto3_(r.MotivoUrgencia)
    }))
  };
}

function versionUGM() {
  return UGM3.VERSION;
}

/* ==========================================================
   UAP · CONSULTA POR DUI (escáner)
   Dice si el paciente de hoy es EXTRAORDINARIO (Reportería
   lo agendó o reagendó para fin de semana), ORDINARIO (cita
   previa de la base HIS) o SOBRECUPO, con nombre, DUI y
   especialidad. No permite marcar ni modificar nada.
   ========================================================== */

// ---------------------------------------------------------
// Para que el escaneo sea instantáneo, el equipo de UAP
// descarga el "paquete" de los días cercanos (paqueteUAP) y
// busca el DUI en el propio equipo. Cada hoja se lee una sola
// vez por paquete, sin importar cuántos días incluya.
// ---------------------------------------------------------

// Base de Reportería de varios días con una sola lectura.
function pacientesFechas3_(fechas) {
  const estado=estadoPacientes3_();
  const out={};
  if (!estado.conFechas) return out;

  const cache=CacheService.getScriptCache();
  const faltan=[];

  fechas.forEach(f => {
    if (!estado.fechas[f]) { out[f]=[]; return; }
    try {
      const g=cache.get('PACDIA_'+estado.version.slice(0,12)+'_'+f);
      if (g) { out[f]=JSON.parse(g); return; }
    } catch(e) {}
    faltan.push(f);
  });

  if (!faltan.length) return out;

  const sh=base3_().getSheetByName('PACIENTES');
  faltan.forEach(f => { out[f]=[]; });
  if (!sh || sh.getLastRow()<2) return out;

  const c=columnasPacientes3_(sh);
  if (c.fecha<0) return out;

  const v=(r,i)=>i>=0 ? texto3_(r[i]) : '';
  const quiero=new Set(faltan);

  sh.getRange(2,1,sh.getLastRow()-1,c.ancho).getValues().forEach(r => {
    const f=v(r,c.fecha);
    if (!quiero.has(f)) return;
    out[f].push({
      nombre:v(r,c.nombre),
      documento:v(r,c.documento),
      original:v(r,c.original),
      especialidad:v(r,c.especialidad).split(' | ')[0] || '',
      hora:v(r,c.hora),
      aumento:v(r,c.aumento)==='SI',
      observacion:v(r,c.observacion),
      expediente:v(r,c.expediente).replace(/\.0+$/,''),
      correlativo:v(r,c.correlativo)
    });
  });

  faltan.forEach(f => {
    try {
      const json=JSON.stringify(out[f]);
      if (json.length<95000) {
        cache.put('PACDIA_'+estado.version.slice(0,12)+'_'+f,json,300);
      }
    } catch(e) {}
  });

  return out;
}

// Índice de un día por DUI, con las hojas ya leídas.
function indiceDia3_(fecha, d) {
  const consultas=consultasDelDia3_(
    fecha,d.agenda,d.solicitudes,d.asistencia,d.pacientes[fecha]||[]
  );

  // Claves: el DUI ("00000000-0") y el expediente ("E:10553-25").
  const p={};
  consultas.forEach(c => {
    const dui=dui3_(c.documento);
    const exp=expUAP3_(c.expediente);
    if (!dui && !exp) return;
    const tipo=c.origen==='SOBRECUPO' ? 'SOBRECUPO'
      : c.extraordinario ? 'EXTRAORDINARIO' : 'ORDINARIO';
    const fila=[
      texto3_(c.paciente),texto3_(c.especialidad),
      texto3_(c.horaCita),tipo,texto3_(c.motivoExtra),
      texto3_(c.asistencia),texto3_(c.marcadoPor),
      horaUAP3_(c.hora),dui,texto3_(c.expediente)
    ];
    if (dui) (p[dui]=p[dui]||[]).push(fila);
    if (exp) (p['E:'+exp]=p['E:'+exp]||[]).push(fila);
  });

  return {t:consultas.length,p};
}

// Lee una vez las hojas que necesita UAP.
function datosUAP3_(fechas) {
  return {
    agenda:leer3_('AGENDA_EXCEL'),
    solicitudes:leer3_('SOBRECUPOS'),
    asistencia:leer3_('ASISTENCIA'),
    pacientes:pacientesFechas3_(fechas)
  };
}

// Expediente normalizado: "10553-25", "202617010" (sin ".0").
function expUAP3_(v) {
  const t=texto3_(v).replace(/\.0+$/,'').toUpperCase().replace(/\s+/g,'');
  return /^[0-9A-Z][0-9A-Z-]{3,}$/.test(t) ? t : '';
}

// Qué trae la lectura: DUI (9 dígitos), expediente, o nada
// útil ("SIN_EXPEDIENTE" que imprimen algunos tickets).
function lecturaUAP3_(lectura) {
  const t=texto3_(lectura).toUpperCase().replace(/\s+/g,'');
  if (!t || /SIN_?EXPEDIENTE|SINEXPEDIENTE/.test(t)) return {sinDato:true};
  const claves=[];
  if (/^\d{8}-?\d$/.test(t)) claves.push(dui3_(t.replace('-','')));
  const exp=expUAP3_(t);
  if (exp) claves.push('E:'+exp);
  return {claves,texto:t};
}

// Hora local (El Salvador) de la marca de asistencia.
function horaUAP3_(v) {
  const t=v instanceof Date ? v : new Date(texto3_(v));
  return isNaN(t.getTime()) ? '' : Utilities.formatDate(t,UGM3.TZ,'HH:mm');
}

// Días con citas cargadas (agenda del HIS o base de Reportería)
// y días con asistencia marcada en los últimos 30 días.
function fechasUAP3_(asistencia) {
  const f={};
  const ag=estadoAgenda3_().porFecha || {};
  const pac=estadoPacientes3_().fechas || {};
  Object.keys(ag).forEach(d => { f[d]=Math.max(f[d]||0,Number(ag[d])||0); });
  Object.keys(pac).forEach(d => {
    const n=Number((pac[d]||{}).total)||0;
    f[d]=Math.max(f[d]||0,n);
  });
  const desde=Utilities.formatDate(
    new Date(Date.now()-30*86400000),UGM3.TZ,'yyyy-MM-dd'
  );
  (asistencia || leer3_('ASISTENCIA')).forEach(a => {
    const d=texto3_(a.Fecha);
    if (d>=desde) f[d]=Math.max(f[d]||0,1);
  });
  return Object.keys(f).filter(d=>/^\d{4}-\d{2}-\d{2}$/.test(d)).sort()
    .map(d=>({fecha:d,total:f[d]}));
}

function fechaUAP3_(valor) {
  try {
    return valor ? fecha3_(valor) : hoy3_();
  } catch(e) {
    return hoy3_();
  }
}

function permisoUAP3_(token) {
  const s=validar3_(token, true);
  if (!['UAP','SUPERVISOR','AUTORIZADO'].includes(s.perfil)) {
    throw Error('Sin permiso para la consulta de UAP.');
  }
  return s;
}

// Días que viajan al equipo: hoy, la fecha elegida y los días
// con citas más cercanos a hoy (hasta 8, de -14 a +14 días).
function diasPaqueteUAP3_(lista, hoy, fecha) {
  const ms=d=>Date.parse(d+'T12:00:00Z');
  const h=ms(hoy);
  const cerca=lista.map(x=>x.fecha)
    .filter(d=>Math.abs(ms(d)-h)<=14*86400000)
    .sort((a,b)=>Math.abs(ms(a)-h)-Math.abs(ms(b)-h));
  return [...new Set([hoy,fecha].concat(cerca))].slice(0,8);
}

function paqueteUAP(token, fechaElegida) {
  const s=permisoUAP3_(token);
  const hoy=hoy3_();
  const fecha=fechaUAP3_(fechaElegida);

  const asistencia=leer3_('ASISTENCIA');
  const lista=fechasUAP3_(asistencia);
  const dias=diasPaqueteUAP3_(lista, hoy, fecha);

  const d={
    agenda:leer3_('AGENDA_EXCEL'),
    solicitudes:leer3_('SOBRECUPOS'),
    asistencia,
    pacientes:pacientesFechas3_(dias)
  };

  const indices={};
  dias.forEach(f => { indices[f]=indiceDia3_(f, d); });

  actividad3_(s,'UAP · Listo para escanear');

  return {
    ok:true,
    hoy,
    fecha,
    fechas:lista.filter(f=>f.fecha>=hoy).slice(0,6),
    dias:indices,
    generado:Date.now(),
    version:UGM3.VERSION
  };
}

// Respaldo: consulta en el servidor (si el equipo aún no tiene
// el paquete). Lee cada hoja una sola vez.
function consultarUAP(token, lectura, fechaElegida) {
  const s=permisoUAP3_(token);

  // El QR del ticket trae el expediente; el DUI también se acepta.
  const lect=lecturaUAP3_(lectura);
  if (lect.sinDato) return {ok:false,resultado:'SIN_DATO'};
  if (!lect.claves.length) {
    return {ok:false,resultado:'LECTURA',lectura:texto3_(lectura).slice(0,30)};
  }
  const dui=lect.texto;

  const hoy=hoy3_();
  const fecha=fechaUAP3_(fechaElegida);
  const asistencia=leer3_('ASISTENCIA');
  const lista=fechasUAP3_(asistencia).filter(f=>f.fecha!==fecha).slice(-12);
  const todas=[fecha].concat(lista.map(f=>f.fecha));
  const d={
    agenda:leer3_('AGENDA_EXCEL'),
    solicitudes:leer3_('SOBRECUPOS'),
    asistencia,
    pacientes:pacientesFechas3_(todas)
  };

  const aCitas=(idx)=>{
    const vistas=new Set();
    const out=[];
    lect.claves.forEach(k => (idx.p[k]||[]).forEach(c => {
      const id=c.join('|');
      if (vistas.has(id)) return;
      vistas.add(id);
      out.push({paciente:c[0],especialidad:c[1],hora:c[2],tipo:c[3],motivo:c[4],
        asistencia:c[5]||'',marcadoPor:c[6]||'',horaMarca:c[7]||'',
        dui:c[8]||'',expediente:c[9]||''});
    }));
    return out;
  };

  const indice=indiceDia3_(fecha, d);
  const citas=aCitas(indice);
  const otras=[];

  if (!citas.length) {
    lista.forEach(f => {
      const c=aCitas(indiceDia3_(f.fecha, d));
      if (c.length) otras.push({fecha:f.fecha,pasada:f.fecha<hoy,citas:c});
    });
  }

  const tipos=[...new Set(citas.map(c=>c.tipo))];
  const resultado=citas.length
    ? (tipos.length===1 ? tipos[0] : 'MIXTO')
    : otras.length ? 'OTRO_DIA' : 'NO_ENCONTRADO';

  const primera=citas[0] || (otras[0] && otras[0].citas[0]);
  registrarUAP(token,[{
    dui,fecha,resultado,
    paciente:primera ? primera.paciente : '',
    especialidad:primera ? primera.especialidad : '',
    otras:otras.map(o=>o.fecha)
  }], s);

  return {ok:true,dui,fecha,hoy,resultado,citas,otras,hayAgenda:indice.t>0};
}

// Compatibilidad con pantallas UAP anteriores.
function latidoUAP(token, fechaElegida) {
  return paqueteUAP(token, fechaElegida);
}

// Auditoría de escaneos por lotes (el equipo los envía en
// segundo plano para no hacer esperar a la ventanilla).
function registrarUAP(token, lote, sesion) {
  const s=sesion || permisoUAP3_(token);
  const filas=(Array.isArray(lote) ? lote : []).slice(0,80).map(e => {
    const detalle=[
      texto3_(e.resultado),texto3_(e.fecha),
      texto3_(e.paciente),texto3_(e.especialidad),
      (e.otras||[]).length ? 'cita el '+e.otras.join(', ') : ''
    ].filter(Boolean).join(' · ');
    const fila={
      FechaHora:e.ts ? new Date(Number(e.ts)).toISOString() : ahora3_(),
      Usuario:s.nombre,
      Perfil:s.perfil,
      Accion:'UAP_ESCANEO',
      Referencia:texto3_(e.dui).slice(0,12),
      Detalle:seguro3_(detalle.slice(0,350))
    };
    return UGM3.HEADERS.AUDITORIA.map(campo =>
      fila[campo]===undefined ? '' : fila[campo]);
  });

  if (filas.length) {
    const sh=hoja3_('AUDITORIA');
    const lock=LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      sh.getRange(sh.getLastRow()+1,1,filas.length,filas[0].length)
        .setValues(filas);
    } finally {
      lock.releaseLock();
    }
  }

  if (!sesion) actividad3_(s,'UAP · Listo para escanear');
  return {ok:true,n:filas.length};
}


/**
 * Cierra de una vez las sesiones sin actividad por más de
 * "minutos" (sesiones pegadas), menos la del supervisor
 * que lo pide.
 */
function cerrarSesionesInactivas(token,minutos) {
  const s=supervisor3_(token);
  const limite=Math.max(5,Number(minutos)||30)*60000;
  const ahora=Date.now();
  let cerradas=0;

  const lock=LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sh=hoja3_('PERSONAL');
    const m=sh.getDataRange().getValues();
    const cab=m[0].map(String);
    const iS=cab.indexOf('Sesion');
    const iU=cab.indexOf('UltimaActividad');
    const iA=cab.indexOf('Actividad');
    const iE=cab.indexOf('Estado');

    for (let i=1;i<m.length;i++) {
      if (String(m[i][iE])!=='CONECTADO') continue;
      if (String(m[i][iS])===s.id) continue;

      const ultima=m[i][iU] instanceof Date
        ? m[i][iU].getTime() : Date.parse(m[i][iU]);

      if (!isNaN(ultima) && ahora-ultima<limite) continue;

      sh.getRange(i+1,iA+1).setValue('Cerrada por inactividad');
      sh.getRange(i+1,iE+1).setValue('DESCONECTADO');
      cerradas++;
    }

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  try {
    auditar3_(s.nombre,s.perfil,'SESION_CERRADA_SUPERVISOR',
      'INACTIVAS',cerradas+' sesiones inactivas');
  } catch(e) {
    console.error(e.message);
  }

  return {ok:true,cerradas};
}

function cerrarSesionRemota(token,sesion) {
  const s=supervisor3_(token);

  const ok=modificarSesion3_(
    {id:texto3_(sesion)},'Cerrada por supervisión',true
  );

  if (!ok) throw Error('Sesión no encontrada.');

  try {
    auditar3_(
      s.nombre,s.perfil,'SESION_CERRADA_SUPERVISOR',
      texto3_(sesion),''
    );
  } catch(e) {
    console.error(e.message);
  }

  return {ok:true};
}

function verTicket(token,id,cual) {
  supervisor3_(token);

  const r=leer3_('SOBRECUPOS').find(
    x=>String(x.ID)===String(id)
  );

  if (!r) throw Error('Solicitud inexistente.');

  const archivoId=cual==='reverso'
    ? r.TicketReverso
    : cual==='comprobante'
      ? r.ComprobanteID
      : (r.TicketFrente || r.ComprobanteID);

  if (!archivoId) return {vacio:true};

  const blob=DriveApp.getFileById(archivoId).getBlob();

  return {
    nombre:blob.getName(),
    mime:blob.getContentType(),
    base64:Utilities.base64Encode(blob.getBytes())
  };
}

/* ==========================================================
   LÍMITES POR ESPECIALIDAD (SUPERVISIÓN)
   ========================================================== */

/**
 * Cambia el límite de cupos de una especialidad.
 * alcance FECHA: solo esa fecha (disponibilidad médica
 *   del día). limite vacío = volver al general.
 * alcance GENERAL: todos los sábados o domingos.
 *   limite vacío = especialidad no habilitada ese día.
 */
function actualizarLimite(token,datos) {
  const s=validar3_(token);

  if (s.perfil!=='SUPERVISOR') {
    throw Error('Solo el supervisor puede cambiar límites.');
  }

  datos=datos || {};

  const especialidad=especialidad3_(datos.especialidad) ||
    texto3_(datos.especialidad).slice(0,100);

  if (!especialidad) throw Error('Especialidad inválida.');

  const vacio=texto3_(datos.limite)==='';
  const limite=vacio ? '' : Number(datos.limite);

  if (!vacio && (!Number.isInteger(limite) ||
      limite<0 || limite>200)) {
    throw Error('El límite debe ser un número entre 0 y 200.');
  }

  let detalle;

  const lock=LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    if (datos.alcance==='GENERAL') {
      const dia=norm3_(datos.dia);

      if (!['SABADO','DOMINGO'].includes(dia)) {
        throw Error('Día inválido.');
      }

      const sh=hoja3_('DISPONIBILIDAD');
      const ultima=sh.getLastRow();
      const filas=ultima>1
        ? sh.getRange(2,1,ultima-1,2).getValues() : [];

      const i=filas.findIndex(r =>
        norm3_(r[0])===dia &&
        norm3_(r[1])===norm3_(especialidad)
      );

      sh.getRange(i>=0 ? i+2 : ultima+1,1,1,3).setValues([[
        dia,especialidad,vacio ? 'N/A' : limite
      ]]);

      detalle=especialidad+' · '+dia+' (general): '+
        (vacio ? 'N/A' : limite);

    } else {
      const fecha=fecha3_(datos.fecha);

      if (dia3_(fecha)==='SEMANA') {
        throw Error('Los límites aplican a sábado y domingo.');
      }

      const sh=hoja3_('LIMITES_FECHA');
      const ultima=sh.getLastRow();
      const filas=ultima>1
        ? sh.getRange(2,1,ultima-1,2).getValues() : [];

      const i=filas.findIndex(r =>
        texto3_(r[0] instanceof Date
          ? Utilities.formatDate(r[0],UGM3.TZ,'yyyy-MM-dd')
          : r[0])===fecha &&
        norm3_(r[1])===norm3_(especialidad)
      );

      if (vacio) {
        if (i>=0) sh.deleteRow(i+2);
      } else {
        sh.getRange(i>=0 ? i+2 : ultima+1,1,1,5).setValues([[
          fecha,especialidad,limite,s.nombre,ahora3_()
        ]]);
      }

      detalle=especialidad+' · '+fecha+': '+
        (vacio ? 'vuelve al general' : limite);
    }

    SpreadsheetApp.flush();

  } finally {
    lock.releaseLock();
  }

  try {
    auditar3_(s.nombre,s.perfil,'LIMITE_CAMBIADO','',detalle);
  } catch(e) {
    console.error(e.message);
  }

  return {ok:true,detalle};
}

function limitesGenerales3_() {
  const mapa={};

  leer3_('DISPONIBILIDAD').forEach(r => {
    const esp=texto3_(r.Especialidad);
    const dia=norm3_(r.Dia);
    if (!esp || !['SABADO','DOMINGO'].includes(dia)) return;
    const x=mapa[esp] || (mapa[esp]={especialidad:esp});
    x[dia]=numeroLimite3_(r.Limite);
  });

  catalogo3_().forEach(e => {
    if (!mapa[e]) mapa[e]={especialidad:e};
  });

  return Object.values(mapa).map(x=>({
    especialidad:x.especialidad,
    sabado:x.SABADO===undefined ? null : x.SABADO,
    domingo:x.DOMINGO===undefined ? null : x.DOMINGO
  }));
}

/* ==========================================================
   BANDEJA DE SOLICITUDES SIN VERIFICAR
   ========================================================== */

function bandeja3_(solicitudes) {
  const ahora=Date.now();
  const hoy=hoy3_();
  const limiteFecha=Utilities.formatDate(
    new Date(ahora+UGM3.URGENTE_DIAS*86400000),
    UGM3.TZ,'yyyy-MM-dd'
  );

  const sinVerificar=solicitudes.filter(r =>
    norm3_(r.Estado)==='PENDIENTE' || !texto3_(r.Estado)
  );

  const urgentes=sinVerificar.filter(r => {
    if (texto3_(r.Urgencia)==='SI') return true;
    const f=texto3_(r.FechaSobrecupo);
    const registro=Date.parse(r.FechaRegistro);
    return (f && f<=limiteFecha) ||
      (!isNaN(registro) &&
       ahora-registro>UGM3.URGENTE_HORAS*3600000);
  });

  const porFecha={};

  sinVerificar.forEach(r => {
    const f=texto3_(r.FechaSobrecupo);
    porFecha[f]=(porFecha[f]||0)+1;
  });

  return {
    total:sinVerificar.length,
    urgentes:urgentes.length,
    ids:sinVerificar.map(r=>String(r.ID)),
    urgentesIds:urgentes.map(r=>String(r.ID)),
    cupoUrgencia:sinVerificar.filter(r=>
      texto3_(r.Urgencia)==='SI'
    ).length,
    vencidas:sinVerificar.filter(r=>
      texto3_(r.FechaSobrecupo)<hoy
    ).length,
    porFecha:Object.keys(porFecha).sort().map(f=>({
      fecha:f,cantidad:porFecha[f]
    }))
  };
}
