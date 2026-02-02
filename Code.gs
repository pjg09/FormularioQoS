function doGet() {
  return HtmlService.createHtmlOutputFromFile("form")
    .setTitle("Recibos de transporte");
}

function obtenerListas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ---------- TECNICOS ----------
  const tecnicos = ss
    .getSheetByName("TECNICOS")
    .getRange("A2:A")
    .getValues()
    .flat()
    .filter(String)
    .map(t => normalizarTexto(t));

  // ---------- LUGARES ----------
  const lugaresSheet = ss.getSheetByName("LUGARES");
  const lugaresData = lugaresSheet
    .getRange("A2:B")
    .getValues()
    .filter(r => r[0]);

  const lugares = [];
  const municipioPorLugar = {};

  lugaresData.forEach(([lugar, municipio]) => {
    const lugarNorm = normalizarTexto(lugar);
    lugares.push(lugarNorm);
    if (municipio) {
      municipioPorLugar[lugarNorm] = normalizarTexto(municipio);
    }
  });

  // ---------- SEDES ----------
  const sedesSheet = ss.getSheetByName("SEDES");
  const sedesData = sedesSheet
    .getRange("A2:C")
    .getValues()
    .filter(r => r[0] && r[1]);

  const sedesPorOrigen = {};
  const municipioPorSede = {};

  sedesData.forEach(([origen, sede, municipio]) => {
    const origenNorm = normalizarTexto(origen);
    const sedeNorm = normalizarTexto(sede);

    if (!sedesPorOrigen[origenNorm]) {
      sedesPorOrigen[origenNorm] = [];
    }
    sedesPorOrigen[origenNorm].push(sedeNorm);

    if (municipio) {
      municipioPorSede[sedeNorm] = normalizarTexto(municipio);
    }
  });

  // ---------- MUNICIPIOS ----------
  const municipiosSheet = ss.getSheetByName("MUNICIPIOS");
  const municipiosData = municipiosSheet
    .getRange("A2:B")
    .getValues()
    .filter(r => r[0] && r[1]);

  const municipios = [];
  const zonaPorMunicipio = {};

  municipiosData.forEach(([municipio, zona]) => {
    const municipioNorm = normalizarTexto(municipio);
    municipios.push(municipioNorm);
    zonaPorMunicipio[municipioNorm] = normalizarTexto(zona);
  });

  // ---------- TARIFAS ----------
  const tarifasSheet = ss.getSheetByName("TARIFAS");
  const tarifasData = tarifasSheet
    .getRange("A2:C")
    .getValues()
    .filter(r => r[0] && r[1]);

  const tarifas = {};

  tarifasData.forEach(([origen, destino, tarifa]) => {
    const origenNorm = normalizarTexto(origen);
    const destinoNorm = normalizarTexto(destino);
    const key = `${origenNorm}|${destinoNorm}`;
    tarifas[key] = Number(tarifa) || 0;
  });

  // ---------- RETURN FINAL ----------
  return {
    tecnicos,
    lugares,
    sedesPorOrigen,
    municipios,
    zonaPorMunicipio,
    municipioPorLugar,
    municipioPorSede,
    tarifas
  };
}

// Función auxiliar para normalizar
function normalizarTexto(texto) {
  return texto
    .toString()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function guardar(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("REGISTROS");

  datos.desplazamientos.forEach(d => {
    hoja.appendRow([
      new Date(),            // Fecha de envio de formulario
      datos.fecha,           // Fecha desplazamiento
      datos.tecnico,         // Nombre y Cédula
      d.origen,              // Lugar origen
      d.sedeOrigen || "",    // Sede origen
      d.destino,             // Lugar destino
      d.sedeDestino || "",   // Sede destino
      d.cantidadBuses,       // Cantidad de buses
      d.valor                // Valor
    ]);
  });

  return "Recibos de transporte registrados correctamente.";
}