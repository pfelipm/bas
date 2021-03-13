/**
 * Importa intervalos de datos idénticos procedentes de todas
 * las hdc que se encuentran en la carpeta indicada
 * y los consolida en la actual.
 * No realiza ningún control de errores.
 * El orden de importación no está garantizado.
 * Demo: https://drive.google.com/drive/folders/1BZNT5TNcOpKaP5Hy3BuvVZhT_FhtrWjr?usp=sharing
 */

function consolidar() {

  const ID_CARPETA_ORIGEN = '1bqlCmxWaL-LCNb6T7vOWNtPky3bK99K-';
  const RANGO_ORIGEN = 'Hoja 1!A2:E';
  const CELDA_DESTINO = 'Hoja 1!A2';

  // Obtener referencias a las hdc dentro de la carpeta
  const hdcsOrigen = DriveApp.getFolderById(ID_CARPETA_ORIGEN).getFilesByType(MimeType.GOOGLE_SHEETS);

  // Lista de IDs de las hdc halladas
  const idHdcs = [];

  // Obtener todos los IDs por medio del iterador
  while (hdcsOrigen.hasNext()) {
    idHdcs.push(hdcsOrigen.next().getId());
  }

  // Abrir cada una de las hdc identificadas e importar datos de manera consolidada
  let datosConsolidados = [];
  idHdcs.forEach(hdc => {
    let datos = SpreadsheetApp.openById(hdc).getRange(RANGO_ORIGEN).getValues();

    // Opcional: elimina filas vacías del intervalo de cada HdC
    datos = datos.filter(fila => fila.some(celda => celda != ''));

    // Consolidar datos
    datosConsolidados = [...datosConsolidados, ...datos];
  });

  // Adaptar dimensiones del intervalo destino a los datos a escribir
  const rangoDestino = SpreadsheetApp.getActive().getRange(CELDA_DESTINO).offset(0, 0, datosConsolidados.length, datosConsolidados[0].length);

  // Escribir datos importados a partir de celda destino
  rangoDestino.setValues(datosConsolidados);

}
