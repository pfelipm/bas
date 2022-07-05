/**
 * Este script importa el contenido de los dos archivos
 * csv indicados en la hoja "Importación", sustituyendo
 * el contenido de las hojas destino o anexando datos.
 * 
 * Se trata de una simple demostración de lo sencillo
 * que resulta acceder al contenido de archivos csv
 * almacenados en Google Drive usando Apps Script.

 * Demo: https://drive.google.com/drive/folders/1QnLKXh5KWSUzzg92hpBYjIviee9N-W3l?usp=sharing
 * 
 * BAS#003 Copyright (C) 2022 Pablo Felip (@pfelipm) · Se distribuye bajo licencia MIT.
 * 
 * @OnlyCurrentDoc
 */

function importarCsv() {

  // Constantes de parametrización del script
  const AJUSTES = {
    hoja: 'Importación',
    nombre1: 'B3',
    nombre2: 'E3',
    hojaDestino1: 'B6',
    hojaDestino2: 'E6',
    anexar: 'B8',
    resultado: 'B12'
  };

  //  Hoja de cálculo y pestaña de ajustes
  const hdc = SpreadsheetApp.getActive()
  const hoja = hdc.getSheetByName(AJUSTES.hoja);

  // Señalizar inicio del proceso
  hoja.getRange(AJUSTES.resultado).setValue('🟠 Importando archivos csv...');
  let resultado = '🔴 No se ha podido realizar la importación';

  // Trata de abrir los archivos csv indicados por el usuario
  const carpeta = DriveApp.getFileById(hdc.getId()).getParents().next();
  const csv1 = carpeta.getFilesByName(hoja.getRange(AJUSTES.nombre1).getValue() + '.csv');
  const csv2 = carpeta.getFilesByName(hoja.getRange(AJUSTES.nombre2).getValue() + '.csv');

  // ¿Existen ambos archivos?
  if (csv1.hasNext() && csv2.hasNext()) {

    // Drive File → Blob → String → String[][] 
    // Espera que el delimitados sea un coma [,], en caso contrario usar
    // parseCsv(csv, delimiter)
    // https://developers.google.com/apps-script/reference/utilities/utilities#parsecsvcsv,-delimiter
    const datos1 = Utilities.parseCsv(csv1.next().getBlob().getDataAsString());
    const datos2 = Utilities.parseCsv(csv2.next().getBlob().getDataAsString());

    // Obtener hojas destino
    const hojaDestino1 = hdc.getSheetByName(hoja.getRange(AJUSTES.hojaDestino1).getValue());
    const hojaDestino2 = hdc.getSheetByName(hoja.getRange(AJUSTES.hojaDestino2).getValue());
    
    // Anexamos o sobreescribimos datos según ajuste en hoja "Importación"
    const anexar = hoja.getRange(AJUSTES.anexar).getValue();
    if (anexar) {
      hojaDestino1.getRange(hojaDestino1.getLastRow() + 1, 1, datos1.length - 1, datos1[0].length)
        .setValues(datos1.slice(1));
      hojaDestino2.getRange(hojaDestino2.getLastRow() + 1, 1, datos1.length - 1, datos2[0].length)
        .setValues(datos2.slice(1));
    } else {
      hojaDestino1.clearContents()
        .getRange(1, 1, datos1.length, datos1[0].length).setValues(datos1);
      hojaDestino2.clearContents()
        .getRange(1, 1, datos2.length, datos2[0].length).setValues(datos2);
    }

    // Si llegamos aquí es que todo ha ido aparentemente bien
    resultado = '🟢 Importación de datos finalizada';

  }

  // Señalizar fin/resultado del proceso
  hoja.getRange(AJUSTES.resultado).setValue(resultado);

}
