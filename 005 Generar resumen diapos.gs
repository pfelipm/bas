/**
 * Genera un tabla de miniaturas de diapositivas que incluye los comentarios del presentador
 * a partir de una presentación de Google en un documento de Google Docs.
 *
 * Demo: https://drive.google.com/drive/folders/1Ui5QZRRUb0kkpzTTyNzjQzsiI9SjPWqv
 * 
 * BAS#005 Copyright (C) Pablo Felip (@pfelipm) · Se distribuye bajo licencia MIT.
 */

/**
 * Añade menú personalizado
 */
function onOpen() {
 DocumentApp.getUi().createMenu('BAS#005')
    .addItem('Generar resumen de presentación', 'generarResumen')
    .addToUi();
}

/**
 * Genera tabla de miniaturas a partir del ID de la presentación
 */
function generarResumen() {
  
  const ui = DocumentApp.getUi();
  let idPresentacion;
  
  try {

    do {
      
      const respuesta = DocumentApp.getUi().prompt('❔ BAS#005', 'Introduce ID de la presentación:', ui.ButtonSet.OK_CANCEL);
      if (respuesta.getSelectedButton() == ui.Button.CANCEL) throw 'Generación cancelada.';
      // Lo apropiado sería usar una expresión regular para extraer el ID a partir del URL, pero eso lo dejaremos para otro BAS...
      idPresentacion = respuesta.getResponseText();

    } while (!idPresentacion);

    const contenido = DocumentApp.getActiveDocument().getBody();
    const presentacion = SlidesApp.openById(idPresentacion);
    const diapos = presentacion.getSlides();
    const diaposId = [];
    const anchuraDiapo = presentacion.getPageWidth();
    const alturaDiapo = presentacion.getPageHeight();
    const anchuraMiniatura = 285;
    const alturaMiniatura = alturaDiapo * anchuraMiniatura / anchuraDiapo;

    // 👇 Necesario para que access_token adquiera el scope necesario (https://www.googleapis.com/auth/drive.readonly)
    // DriveApp.getFileById(idPresentacion);
    // Sí, increiblemente basta con que esté comentado para que se incluya su scope en el cuadro de diálogo de autorización

    // Construye el URL "mágico" de generación de PNG (ver apartado 2.4 del BAS#002)
    const url = `https://docs.google.com/presentation/d/${idPresentacion}/export/png?access_token=${ScriptApp.getOAuthToken()}`; 

    // Ahora obtendremos cada diapo como imagen PNG... ¡en paralelo!
    const peticionesPng = diapos.map(diapo => {

      // Generamos también un array con todos los ID de las diapositivas
      const diapoId = diapo.getObjectId();
      diaposId.push(diapoId)

      // URL para obtener la imagen de la diapositiva en png
      return {'url': `${url}&pageid=${diapoId}`}
      
    });

    const imagenesPng = UrlFetchApp.fetchAll(peticionesPng);
    const numDiapos = imagenesPng.length;

    // ¿Sabías que una presentación puede existir pero NO tener ninguna diapositiva?
    if (numDiapos == 0) throw 'La presentación seleccionada no contiene dispositivas.'

    // Tenemos diapos que obtener, SOLO ahora ahora borraremos el cuerpo del documento para generar nuevas miniaturas
    contenido.clear();

    // Esta vez además mediremos el tiempo de ejecución
    const t1 = new Date();

    // Inserta cada imagen en una tabla:
    // |----------------------------------------------|
    // | Diapositiva nº n de m  |  Notas diapositiva  |
    // |----------------------------------------------|
    // |   Imagen en miniatura  |  Notas presentador  |
    // |----------------------------------------------|
    // Estructura de un DOC https://developers.google.com/apps-script/guides/docs?hl=en#structure_of_a_document

    imagenesPng.forEach((imagen, indiceDiapo) => {
      
      // Si no estamos en la 1ª página añade un párrafo para que todas las tablas comiencen en la misma posición
      // dado que la 1ª página siempre contiene una línea en blanco
      if ((indiceDiapo + 1) % 3 == 1 && indiceDiapo > 2) contenido.appendParagraph('');

      // Construye la tabla para cada diapositiva
      const tabla = contenido.appendTable([[`Diapositiva nº ${indiceDiapo + 1} de ${numDiapos}`, 'Notas de la diapositiva']]);
      const fila = tabla.appendTableRow();
      fila.appendTableCell().appendImage(imagen.getBlob()).setWidth(anchuraMiniatura).setHeight(alturaMiniatura)
        .setLinkUrl(`${presentacion.getUrl()}#slide=id.${diaposId[indiceDiapo]}`);
      fila.appendTableCell().appendParagraph(diapos[indiceDiapo].getNotesPage().getSpeakerNotesShape().getText().asString());

      // Formatea celdas de la tabla (encabezado y bordes)
      const atributosEncabezado = {};
      atributosEncabezado[DocumentApp.Attribute.BOLD] = true;
      atributosEncabezado[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FFFFFF';  // Blanco
      tabla.getRow(0).setAttributes(atributosEncabezado);
      tabla.getRow(0).getCell(0).setBackgroundColor('#4E5D6C'); // Carbón
      tabla.getRow(0).getCell(1).setBackgroundColor('#4E5D6C');
      tabla.setBorderColor('#4E5D6C');

      // Inserta una zona de notas y un salto de página cada 3 diapositivas (y al final)
      // para garantizar que las tablas no queden cortadas

      if ((indiceDiapo + 1) % 3 == 0 || indiceDiapo == numDiapos - 1){

        contenido.appendParagraph('Otras notas:').setSpacingAfter(3).setBold(true);
        // contenido.appendParagraph('').setBold(false);
        contenido.appendPageBreak();
        
      }

    });

    // Inicializa el encabezado del documento
    let encabezado;
    encabezado = DocumentApp.getActiveDocument().getHeader();
    if (encabezado) encabezado.clear();
    else encabezado = DocumentApp.getActiveDocument().addHeader();

    // Construye el encabezado con enlace a la presentación
    encabezado.appendParagraph('Miniaturas de ').setFontSize(11)
      .appendText(presentacion.getName())
      .setLinkUrl(presentacion.getUrl());

    // Añade marca de tiempo para datar el documento generado
    // Es necesario indicar el "locale" porque Session.getActiveUserLocale() ahora mismo no funciona bien >> https://issuetracker.google.com/issues/179563675
    encabezado.appendParagraph(`Generado el ${t1.toLocaleDateString('es')} a las ${t1.toLocaleTimeString('es')}`)
      .setFontSize(8)
      .setSpacingAfter(6);

    const t2 = new Date();

    // Mensaje de fin del proceso (con éxito)
    DocumentApp.getUi().alert('🟢 BAS#005',
                              `${numDiapos} miniatura(s) generada(s) en aproximadamente ${Math.round((t2 - t1)/1000)}".`,
                              ui.ButtonSet.OK);

  } catch (e) {
    DocumentApp.getUi().alert('🔴 BAS#005', typeof e == 'string' ? e : `Error interno: ${e.message}.`, ui.ButtonSet.OK);
  }

}
