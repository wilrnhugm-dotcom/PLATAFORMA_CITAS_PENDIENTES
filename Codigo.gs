
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
 * - 1 sobrecupo por 10 pacientes activos únicos.
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
  TZ: 'America/El_Salvador',

  // Claves exactas: respetar mayúsculas y espacios.
  ROLES: {
    SUPERVISOR: 'SUP 1',
    AUTORIZADO: 'HNRUGM$',
    TECNICO: 'UGM 2026'
  },

  // Sin nuevas programaciones ni asignación de cupos.
  FECHAS_BLOQUEADAS: [
    '2026-12-24', '2026-12-25', '2026-12-26',
    '2026-12-27', '2026-12-31',
    '2027-01-01', '2027-01-02', '2027-01-03'
  ],

  VIGENCIA_MINUTOS: 30,
  ALERTA_MINUTOS: 60,
  MAX_CITAS: 15000,
  MAX_ARCHIVO_MB: 5,
  MAX_PACIENTES: 250000,
  LOTE_PACIENTES: 5000,
  PAGINA_PACIENTES: 20000,
  MAX_RESULTADOS: 20,

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
      'Observacion'
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
      'IDExterno'
    ],

    PACIENTES: [
      'Nombre',
      'Documento',
      'DocumentoOriginal',
      'Expediente',
      'Especialidad',
      'Busqueda'
    ]
  }
};

/* ==========================================================
   APLICACIÓN
   ========================================================== */

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('UGM | Control de Sobrecupos')
    .addMetaTag(
      'viewport',
      'width=device-width,initial-scale=1'
    );
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
  const digitos = texto3_(valor)
    .replace(/\.0+$/, '')
    .replace(/\D/g, '');

  const nueve = digitos.length === 8
    ? '0' + digitos : digitos;

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

function hoja3_(nombre) {
  const sh = base3_().getSheetByName(nombre);

  if (!sh) {
    throw Error('No existe la hoja: ' + nombre);
  }

  return sh;
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

    sh.getRange(1,1,1,7)
      .setBackground('#10456B')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');

    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('E:E').setNumberFormat('@');
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

  return Utilities.base64EncodeWebSafe(
    cuerpo+'|'+firma
  );
}

function validar3_(token) {
  if (!token) throw Error('Inicia sesión.');

  let partes;

  try {
    partes = Utilities.newBlob(
      Utilities.base64DecodeWebSafe(token)
    ).getDataAsString().split('|');
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

  return s;
}

// Cada enlace (?vista=tecnico / ?vista=supervisor)
// solo admite sus perfiles. Sin parámetro, admite todos.
const VISTAS3 = {
  tecnico: ['TECNICO','AUTORIZADO'],
  supervisor: ['SUPERVISOR','AUTORIZADO']
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

  if (VISTAS3[v] && !VISTAS3[v].includes(perfil)) {
    throw Error(
      v === 'tecnico'
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

  return {token,nombre,perfil};
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
  const s = validar3_(token);

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

  const errores = {
    sinFecha:0,
    sinEspecialidad:0,
    estadoNoReconocido:0
  };

  registros.forEach((r,i) => {
    let f;

    try {
      f = fechaAgenda3_(r.fecha);
    } catch(e) {
      errores.sinFecha++;
      return;
    }

    const esp = especialidad3_(r.especialidad);

    if (!esp) {
      errores.sinEspecialidad++;
      return;
    }

    const estado = estadoCita3_(r.estado);

    if (estado === 'OTRA') {
      errores.estadoNoReconocido++;
      return;
    }

    const doc = documento3_(r.documento);
    const paciente = texto3_(r.paciente).slice(0,150);
    const id = texto3_(r.id).slice(0,100);

    if (!doc && !paciente && !id) {
      return;
    }

    fechas.add(f);

    validos.push([
      '', f, esp, estado,
      doc, paciente, id || String(i+1)
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
  if (validos.length < registros.length * 0.5) {
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
        2,1,ultima-1,7
      ).getValues();
      sh.getRange(
        2,1,ultima-1,7
      ).clearContent();
    }

    try {
      for (let i=0;i<validos.length;i+=500) {
        const lote = validos
          .slice(i,i+500)
          .map(fila => fila.map(seguro3_));

        sh.getRange(
          i+2,1,lote.length,7
        ).setValues(lote);
      }

      SpreadsheetApp.flush();

    } catch(error) {
      // Recuperación de la hoja anterior si falla
      // la escritura de la nueva agenda.
      const fin = sh.getLastRow();

      if (fin > 1) {
        sh.getRange(
          2,1,fin-1,7
        ).clearContent();
      }

      if (anterior && anterior.length) {
        for (let i=0;i<anterior.length;i+=500) {
          const lote=anterior.slice(i,i+500);
          sh.getRange(
            i+2,1,lote.length,7
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
    omitidas:registros.length-validos.length,
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

  return {
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
    alerta:
      minutos === null ||
      minutos >= UGM3.ALERTA_MINUTOS ||
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

function limite3_(fecha,especialidad) {
  const dia = dia3_(fecha);

  if (dia === 'SEMANA') return null;

  const encontrado = leer3_('DISPONIBILIDAD')
    .find(r =>
      norm3_(r.Dia) === dia &&
      norm3_(r.Especialidad) ===
        norm3_(especialidad)
    );

  if (!encontrado ||
      texto3_(encontrado.Limite) === '' ||
      norm3_(encontrado.Limite) === 'N/A') {
    return null;
  }

  const n = Number(encontrado.Limite);

  return Number.isInteger(n) && n >= 0
    ? n : null;
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
  const dia=dia3_(fecha);

  if (dia === 'SEMANA') return [];

  const limites=leer3_('DISPONIBILIDAD')
    .filter(r=>norm3_(r.Dia)===dia);

  return limites.map(r => {
    const especialidad=texto3_(r.Especialidad);

    const pacientes=contarAgenda3_(
      agenda,fecha,especialidad
    );

    const limite=limite3_(
      fecha,especialidad
    );

    const base=Math.floor(
      pacientes.activas/10
    );

    const bloqueada=bloqueada3_(fecha);

    const autorizados=limite===null || bloqueada
      ? 0 : Math.min(base,limite);

    const ocupados=ocupados3_(
      solicitudes,fecha,especialidad
    );

    return {
      especialidad,
      activas:pacientes.activas,
      pendientes:pacientes.pendientes,
      limite,
      generados:autorizados,
      ocupados,
      disponibles:Math.max(
        0,autorizados-ocupados
      ),
      habilitada:limite!==null && !bloqueada,
      bloqueada,
      incidencia:!bloqueada && ocupados>autorizados
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

  return datos;
}

/* ==========================================================
   GUARDAR SOBRECUPO
   ========================================================== */

function crearSolicitud(token,datos) {
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

  if (dia3_(fecha)==='SEMANA') {
    fallo3_(
      'DIA_NO_HABILITADO',
      'No está permitido registrar sobrecupos '+
      'de lunes a viernes.'
    );
  }

  if (!especialidad ||
      limite3_(fecha,especialidad)===null) {
    fallo3_(
      'ESPECIALIDAD_NO_HABILITADA',
      'Esta especialidad no tiene sobrecupos '+
      'habilitados para ese día.'
    );
  }

  const lock=LockService.getScriptLock();
  lock.waitLock(30000);

  let id;
  let disponibles;

  try {
    const vigente=estadoAgenda3_();

    if (!vigente.vigente) {
      fallo3_(
        'AGENDA_DESACTUALIZADA',
        'Actualiza la agenda hospitalaria antes '+
        'de registrar el sobrecupo. Se requiere '+
        'una carga válida de los últimos 30 minutos.'
      );
    }

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

    const agenda=leer3_('AGENDA_EXCEL');

    const pacientes=contarAgenda3_(
      agenda,fecha,especialidad
    );

    const limite=limite3_(
      fecha,especialidad
    );

    const generados=Math.min(
      Math.floor(pacientes.activas/10),
      limite
    );

    const ocupados=ocupados3_(
      existentes,fecha,especialidad
    );

    if (ocupados>=generados) {
      fallo3_(
        'SIN_CUPOS',
        'No hay cupos disponibles para '+
        especialidad+' el '+fecha+'. '+
        'Pacientes activos: '+pacientes.activas+
        '. Sobrecupos autorizados: '+generados+
        '. Ocupados: '+ocupados+'.'
      );
    }

    id='UGM-'+Utilities.formatDate(
      new Date(),UGM3.TZ,'yyyyMMdd'
    )+'-'+Utilities.getUuid()
      .slice(0,8).toUpperCase();

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
      Estado:'PENDIENTE'
    });

    SpreadsheetApp.flush();

    disponibles=generados-ocupados-1;

  } finally {
    lock.releaseLock();
  }

  try {
    auditar3_(
      s.nombre,s.perfil,
      'SOBRECUPO_GUARDADO',
      id,especialidad+' / '+fecha
    );
  } catch(e) {
    console.error(e.message);
  }

  return {
    ok:true,
    id,
    disponibles,
    mensaje:'Sobrecupo guardado correctamente.'
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

    if (['REGISTRADO','RECHAZADO','ANULADO']
      .includes(actual)) {
      throw Error('La solicitud ya está cerrada.');
    }

    if (nuevoEstado==='REGISTRADO') {
      const comprobante=texto3_(m[i][12]);

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
    conEspecialidad:c.PAC_ESPECIALIDAD==='1'
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

    return [
      seguro3_(nombre),
      documento,
      seguro3_(original),
      seguro3_(texto3_(r.expediente).slice(0,50)),
      seguro3_(especialidad),
      busqueda3_(nombre)
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
        UGM3_PAC_CARGA_ESP:''
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

    p.setProperties({
      UGM3_PAC_CARGA_INV:String(acumulados),
      UGM3_PAC_CARGA_ESP:tieneEspecialidad ? '1' : ''
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
        PAC_ESPECIALIDAD:tieneEspecialidad ? '1' : '0'
      });

      p.setProperty('UGM3_PAC_CARGA','');

      final={total,invalidos:acumulados};
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
    invalidos:final ? final.invalidos : 0
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
      v(r,c.especialidad)
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
    if (!esp) return;

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
