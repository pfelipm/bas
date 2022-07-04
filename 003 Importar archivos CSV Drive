/**
 * Exporta todas las diapositivas de la presentación como imágenes PNG
 * en una carpeta junto a la propia presentación.
 * No se realiza ningún control de errores.
 * Demo: https://drive.google.com/drive/u/0/folders/1dahbbyFE-3ixc_rHOA8--ufx3idgSGBM
 * 
 * BAS#002 Copyright (C) Pablo Felip (@pfelipm) · Se distribuye bajo licencia MIT.
 */

/**
 * Añadir menú personalizado
 */
function onOpen() {
  SlidesApp.getUi().createMenu('Slides2PNG')
    .addItem('Exportar diapositivas como PNG', 'exportarDiaposPngUrl')
    .addToUi();
}

/* Exporta todas las diapos como png en carpeta de Drive junto a la presentación */
function exportarDiaposPngUrl() {

  // Presentación sobre la que estamos trabajando
  const presentacion = SlidesApp.getActivePresentation();
  const idPresentacion = presentacion.getId();
  // Presentación en Drive
  const presentacionDrive = DriveApp.getFileById(idPresentacion);
  // Carpeta donde se encuentra en la presentación
  const carpeta  = presentacionDrive.getParents().next();
  // Nombre de la carpeta de exportación para los PNG
  const nombreCarpetaExp = `Miniaturas {${idPresentacion}}`; 

  // Si la carpeta de exportación ya existe la eliminamos para evitar duplicados (¡con el mismo nombre!)
  if (carpeta.getFoldersByName(nombreCarpetaExp).hasNext()) {
    carpeta.getFoldersByName(nombreCarpetaExp).next().setTrashed(true);
  }

  // Crear carpeta de exportación
  const carpetaExp = carpeta.createFolder(nombreCarpetaExp);

  // Lista de diapositivas en la presentación
  const diapos = presentacion.getSlides();

  // ¿Cuántos dígitos necesitamos para representar el nº de orden de la imagen exportada?
  const nDigitos = parseInt(diapos.length.toString().length);

  // URL "mágico" para la exportación PNG
  const url = `https://docs.google.com/presentation/d/${idPresentacion}/export/png?access_token=${ScriptApp.getOAuthToken()}`; 

  // Enumerar diapositivas y exportar en formato PNG
  diapos.forEach((diapo, num) => {
  
    // Obtener blob de la diapositiva exportada en png
    const blobDiapo = UrlFetchApp.fetch(`${url}&pageid=${diapo.getObjectId()}`).getBlob();

    // Por fin, creamos imágenes a partir de los blobs obtenidos para cada diapo,
    // nombres precedidos por nº de diapositiva con relleno de 0s por la izquierda
    carpetaExp.createFile(blobDiapo.setName(`Diapositiva ${String(num + 1).padStart(nDigitos, '0')}`));
  
  });

}
