const CARPETA_CDC_ID = "1dAcJ24CwmS1C3_jmwE4AV6KYQ41ZNNah";
const DIAS_ATRAS_A_REVISAR = 15;

function actualizarDiferencias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const datos = ss.getSheetByName("DATOS");
  if (!datos) throw new Error("No existe la pestaña DATOS");

  // IDs ya existentes en DATOS (columna B)
  const ultimaFilaDatos = datos.getLastRow();
  const idsExistentes = new Set();
  if (ultimaFilaDatos > 1) {
    const idsRango = datos.getRange(2, 2, ultimaFilaDatos - 1, 1).getValues();
    idsRango.forEach(fila => {
      const id = String(fila[0]).trim();
      if (id) idsExistentes.add(id);
    });
  }

  // Fecha límite: hoy - DIAS_ATRAS_A_REVISAR
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  const fechaLimite = new Date(hoy);
  fechaLimite.setDate(fechaLimite.getDate() - DIAS_ATRAS_A_REVISAR);

  const carpeta = DriveApp.getFolderById(CARPETA_CDC_ID);
  const archivos = carpeta.getFilesByType(MimeType.GOOGLE_SHEETS);

  const filasNuevas = [];
  let archivosRevisados = 0;
  let ignoradosPorError = 0;

  while (archivos.hasNext()) {
    const archivo = archivos.next();
    const nombre = archivo.getName();

    if (!nombre.startsWith("CDC")) continue;

    const fecha = extraerFechaDeNombre(nombre);
    if (!fecha) continue;

    if (fecha < fechaLimite) continue;

    archivosRevisados++;

    const libroCdc = SpreadsheetApp.openById(archivo.getId());
    const hojaResumen = libroCdc.getSheetByName("Resumen de Cargas");
    if (!hojaResumen) continue;

    const diferencias = extraerDiferencias(hojaResumen);

    diferencias.forEach(d => {
      const idGuia = String(d.guia).trim();
      const estatus = String(d.estatus).trim();

      if (!idGuia) return;

      if (esValorError(idGuia) || esValorError(estatus)) {
        ignoradosPorError++;
        return;
      }

      if (idsExistentes.has(idGuia)) return;

      Logger.log("Archivo: " + nombre + " | Guía: " + idGuia + " | Fecha a guardar: " + fecha.toString());

      filasNuevas.push([fecha, idGuia, estatus, "", "", ""]);
      idsExistentes.add(idGuia);
    });
  }

  if (filasNuevas.length > 0) {
    datos.getRange(datos.getLastRow() + 1, 1, filasNuevas.length, 6)
         .setValues(filasNuevas);
  }

  // ===== Normalizar columna FECHA a valores Date reales antes de ordenar =====
  const filaFinal = datos.getLastRow();
  if (filaFinal > 2) {
    const rangoFechas = datos.getRange(2, 1, filaFinal - 1, 1);
    const valoresFechas = rangoFechas.getValues();

    const fechasNormalizadas = valoresFechas.map(fila => {
      const val = fila[0];

      if (val instanceof Date) return [val];

      // Si es texto tipo "08-07-26" o "08/07/2026", intentar convertir
      const texto = String(val).trim();
      const match = texto.match(/(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})/);
      if (match) {
        let dia = parseInt(match[1], 10);
        let mes = parseInt(match[2], 10);
        let anio = parseInt(match[3], 10);
        if (anio < 100) anio += 2000;
        return [new Date(anio, mes - 1, dia)];
      }

      return [val]; // no se pudo convertir, se deja igual
    });

    rangoFechas.setValues(fechasNormalizadas);

    // Ahora sí, ordenar toda la tabla por FECHA (columna A) ascendente
    datos.getRange(2, 1, filaFinal - 1, datos.getLastColumn())
         .sort({ column: 1, ascending: true });
  }

  if (filasNuevas.length === 0) {
    SpreadsheetApp.getUi().alert(
      "Se revisaron " + archivosRevisados + " archivo(s) de los últimos " +
      DIAS_ATRAS_A_REVISAR + " días. No se encontraron diferencias nuevas." +
      (ignoradosPorError > 0 ? " (" + ignoradosPorError + " ignorada(s) por error #N/A)" : "")
    );
    return;
  }

  SpreadsheetApp.getUi().alert(
    filasNuevas.length + " diferencia(s) nueva(s) agregada(s) (de " +
    archivosRevisados + " archivo(s) revisados)." +
    (ignoradosPorError > 0 ? " " + ignoradosPorError + " ignorada(s) por error #N/A." : "")
  );
}

function esValorError(valor) {
  return /^#(N\/A|REF!|ERROR!|VALUE!|DIV\/0!|NULL!|NUM!|NAME\?)/.test(valor);
}

function extraerFechaDeNombre(nombre) {
  const match = nombre.match(/(\d{2})-(\d{2})-(\d{2})/);
  if (!match) return null;

  const dia = parseInt(match[1], 10);
  const mes = parseInt(match[2], 10);
  const anio = 2000 + parseInt(match[3], 10);

  return new Date(anio, mes - 1, dia);
}

function extraerDiferencias(hoja) {
  const valores = hoja.getDataRange().getValues();
  const resultado = [];

  for (let fila = 0; fila < valores.length; fila++) {
    for (let col = 0; col < valores[fila].length; col++) {

      if (String(valores[fila][col]).trim().toUpperCase() === "GUÍA") {
        const colGuia = col;
        const colEstatus = col + 1;

        for (let f = fila + 1; f < valores.length; f++) {
          const guia = valores[f][colGuia];
          const estatus = valores[f][colEstatus];

          if (!guia) break;

          resultado.push({ guia: guia, estatus: estatus });
        }
      }
    }
  }

  return resultado;
}

function probarFecha() {
  const resultado = extraerFechaDeNombre("CDC 16-07-26");
  Logger.log(resultado);
  Logger.log(resultado.toString());
}
