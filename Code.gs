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
    .filter(String);
  
  // ---------- LUGARES ----------
  const lugaresSheet = ss.getSheetByName("LUGARES");
  const lugaresData = lugaresSheet
    .getRange("A2:B")
    .getValues()
    .filter(r => r[0]);
  
  const lugares = [];
  const municipioPorLugar = {};
  
  lugaresData.forEach(([lugar, municipio]) => {
    lugares.push(lugar);
    if (municipio) {
      municipioPorLugar[lugar] = municipio;
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
    if (!sedesPorOrigen[origen]) {
      sedesPorOrigen[origen] = [];
    }
    sedesPorOrigen[origen].push(sede);
    if (municipio) {
      municipioPorSede[sede] = municipio;
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
    municipios.push(municipio);
    zonaPorMunicipio[municipio] = zona.toUpperCase();
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
    .replace(/[\u0300-\u036f]/g, ""); // Elimina tildes
}

function guardar(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("REGISTROS");
  
  datos.desplazamientos.forEach(d => {
    hoja.appendRow([
      new Date(),          // Fecha de registro
      datos.fecha,         // Fecha
      datos.tecnico,       // Técnico
      d.origen,            // Origen
      d.sedeOrigen || "",  // Sede origen
      d.destino,           // Destino
      d.sedeDestino || "", // Sede destino
      d.valor              // Valor
    ]);
  });
  
  return "Recibos de transporte registrados correctamente.";
}