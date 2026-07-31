/**
 * Backend de DCE Recepción.
 * Este proyecto debe estar vinculado al Google Sheets Base_Datos_DCE_Recepcion.
 */

const HOJAS = Object.freeze({
  ANUNCIOS: "Anuncios",
  CUMPLEANOS: "Cumpleanos",
  EVENTOS: "Eventos",
  CONFIGURACION: "Configuracion"
});

const ZONA_HORARIA_PREDETERMINADA = "America/Monterrey";

function doGet(e) {
  let respuesta;

  try {
    respuesta = obtenerDatos_();
  } catch (error) {
    respuesta = {
      ok: false,
      updatedAt: new Date().toISOString(),
      error: error && error.message ? error.message : String(error)
    };
  }

  const callbackSolicitado = e && e.parameter
    ? String(e.parameter.callback || "")
    : "";
  const callbackValido = /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callbackSolicitado)
    ? callbackSolicitado
    : "";
  const contenido = callbackValido
    ? callbackValido + "(" + JSON.stringify(respuesta) + ");"
    : JSON.stringify(respuesta);

  return ContentService
    .createTextOutput(contenido)
    .setMimeType(
      callbackValido
        ? ContentService.MimeType.JAVASCRIPT
        : ContentService.MimeType.JSON
    );
}

function obtenerDatos_() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  validarHojas_(libro);

  const ahora = new Date();
  const zona = obtenerZonaHoraria_(libro);

  const anuncios = leerObjetos_(libro.getSheetByName(HOJAS.ANUNCIOS))
    .filter(fila => esActivo_(fila.Activo))
    .filter(fila => estaVigente_(fila.FechaInicio, fila.FechaFin, ahora))
    .map(fila => ({
      id: texto_(fila.ID),
      category: texto_(fila.Categoria) || "Aviso",
      title: texto_(fila.Titulo),
      description: texto_(fila.Descripcion),
      date: texto_(fila.FechaTexto) || rangoFechas_(fila.FechaInicio, fila.FechaFin, zona),
      location: texto_(fila.Ubicacion),
      accent: texto_(fila.Color) || "#171717",
      priority: numero_(fila.Prioridad, 99),
      active: true
    }))
    .filter(anuncio => anuncio.title)
    .sort((a, b) => a.priority - b.priority);

  const cumpleanos = leerObjetos_(libro.getSheetByName(HOJAS.CUMPLEANOS))
    .filter(fila => esActivo_(fila.Activo))
    .map(fila => ({
      id: texto_(fila.ID),
      name: texto_(fila.Nombre),
      role: texto_(fila.Puesto),
      type: texto_(fila.Tipo) || "Personal",
      day: numero_(fila.Dia, 0),
      month: numero_(fila.Mes, 0),
      photo: texto_(fila.Foto),
      active: true
    }))
    .filter(persona =>
      persona.name &&
      persona.day >= 1 && persona.day <= 31 &&
      persona.month >= 1 && persona.month <= 12
    );

  const eventos = leerObjetos_(libro.getSheetByName(HOJAS.EVENTOS))
    .filter(fila => esActivo_(fila.Activo))
    .map(fila => ({
      id: texto_(fila.ID),
      title: texto_(fila.Titulo),
      day: numero_(fila.Dia, 0),
      month: numero_(fila.Mes, 0),
      year: numero_(fila.Ano, 0),
      time: horaTexto_(fila.Hora, zona),
      location: texto_(fila.Lugar),
      type: texto_(fila.Tipo) || "Evento",
      active: true
    }))
    .filter(evento =>
      evento.title &&
      esFechaValida_(evento.year, evento.month, evento.day)
    )
    .sort((a, b) =>
      a.year - b.year ||
      a.month - b.month ||
      a.day - b.day ||
      a.time.localeCompare(b.time)
    );

  return {
    ok: true,
    updatedAt: ahora.toISOString(),
    announcements: anuncios,
    birthdays: cumpleanos,
    events: eventos,
    config: leerConfiguracion_(libro.getSheetByName(HOJAS.CONFIGURACION))
  };
}

function validarHojas_(libro) {
  const faltantes = Object.keys(HOJAS)
    .map(clave => HOJAS[clave])
    .filter(nombre => !libro.getSheetByName(nombre));

  if (faltantes.length) {
    throw new Error("Faltan las hojas: " + faltantes.join(", "));
  }
}

function leerObjetos_(hoja) {
  if (!hoja || hoja.getLastRow() < 2) return [];

  const valores = hoja.getDataRange().getValues();
  const encabezados = valores.shift().map(texto_);

  return valores.map(fila =>
    encabezados.reduce((objeto, encabezado, indice) => {
      if (encabezado) objeto[encabezado] = fila[indice];
      return objeto;
    }, {})
  );
}

function leerConfiguracion_(hoja) {
  return leerObjetos_(hoja).reduce((resultado, fila) => {
    const clave = texto_(fila.Clave);
    if (clave) resultado[clave] = texto_(fila.Valor);
    return resultado;
  }, {});
}

function obtenerZonaHoraria_(libro) {
  const configuracion = leerConfiguracion_(
    libro.getSheetByName(HOJAS.CONFIGURACION)
  );
  return configuracion.zona_horaria ||
    Session.getScriptTimeZone() ||
    ZONA_HORARIA_PREDETERMINADA;
}

function esActivo_(valor) {
  const normalizado = texto_(valor).toUpperCase();
  return !["NO", "FALSE", "FALSO", "0", "INACTIVO"].includes(normalizado);
}

function estaVigente_(inicio, fin, ahora) {
  const desde = fecha_(inicio, false);
  const hasta = fecha_(fin, true);
  return (!desde || ahora >= desde) && (!hasta || ahora <= hasta);
}

function fecha_(valor, finDelDia) {
  if (!valor) return null;

  let fecha;
  if (valor instanceof Date) {
    fecha = new Date(valor);
  } else {
    const cadena = texto_(valor);
    const iso = cadena.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
    const local = cadena.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);

    if (iso) {
      fecha = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    } else if (local) {
      fecha = new Date(Number(local[3]), Number(local[2]) - 1, Number(local[1]));
    } else {
      fecha = new Date(cadena);
    }
  }

  if (isNaN(fecha.getTime())) return null;
  fecha.setHours(finDelDia ? 23 : 0, finDelDia ? 59 : 0, finDelDia ? 59 : 0, 0);
  return fecha;
}

function esFechaValida_(ano, mes, dia) {
  if (![ano, mes, dia].every(Number.isFinite)) return false;
  const fecha = new Date(ano, mes - 1, dia);
  return fecha.getFullYear() === ano &&
    fecha.getMonth() === mes - 1 &&
    fecha.getDate() === dia;
}

function rangoFechas_(inicio, fin, zona) {
  const formato = valor => valor
    ? Utilities.formatDate(new Date(valor), zona, "d MMM")
    : "";
  const desde = formato(inicio);
  const hasta = formato(fin);
  return desde && hasta ? desde + " – " + hasta : desde || hasta || "Vigente";
}

function horaTexto_(valor, zona) {
  if (!valor) return "";
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, zona, "HH:mm");
  }
  return texto_(valor);
}

function texto_(valor) {
  return valor === null || valor === undefined ? "" : String(valor).trim();
}

function numero_(valor, respaldo) {
  const numero = Number(valor);
  return Number.isFinite(numero) ? numero : respaldo;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("HMI DCE")
    .addItem("Probar conexión y datos", "probarLectura")
    .addToUi();
}

function probarLectura() {
  const datos = obtenerDatos_();
  SpreadsheetApp.getUi().alert(
    "HMI DCE",
    "Conexión correcta" +
      "\nAnuncios activos: " + datos.announcements.length +
      "\nCumpleaños registrados: " + datos.birthdays.length +
      "\nEventos activos: " + datos.events.length,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function probarRespuestaWeb() {
  const datos = obtenerDatos_();
  Logger.log(JSON.stringify(datos, null, 2));
  return datos;
}
