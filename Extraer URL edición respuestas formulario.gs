/** 
 * Este script extrae los URL de edición de las respuestas recibidas en el formulario
 * indicado por el usuario en la celda PARAMETROS.url, además de otros campos adicionales
 * opcionales. Todas las respuestas deben tener el mismo nº de preguntas que se recuperan
 * como identificación (celda numCampos).
 * 
 * Demo: https://docs.google.com/spreadsheets/d/1kSpZbNiBJWAKJVqxzJ4ATbChsXMGNlh2fW_hDwSpK8A/edit?usp=sharing
 * 
 * BAS#004 Copyright (C) 2022 Pablo Felip (@pfelipm) · Se distribuye bajo licencia MIT.
 * 
 * @OnlyCurrentDoc
 */
function resumirRespuestas() {

  // Constantes de parametrización del script
  const PARAMETROS = {
    filaTabla: 8,
    url: 'B1',
    numCampos: 'B3',
    fechaSiNo: 'B4',
    emailSiNo: 'B5',
    urlSiNo: 'B6'
  };

  // Hoja de cálculo
  const hdc = SpreadsheetApp.getActive();
  const hoja = hdc.getActiveSheet();

  // Leer parámetros
  const numCampos = hoja.getRange(PARAMETROS.numCampos).getValue();
  const fechaSiNo = hoja.getRange(PARAMETROS.fechaSiNo).getValue();
  const emailSiNo = hoja.getRange(PARAMETROS.emailSiNo).getValue();
  // Sí, este parámetros es 'fake', dado que la la hdc no permite desmarcar la casilla, pero ahí queda
  const urlSiNo = hoja.getRange(PARAMETROS.urlSiNo).getValue();

  // Sección principal, que se ejecuta dentro de un bloque en el que
  // se cazarán los errores en tiempo de ejecución.
  try {

    // Acceder al formulario objetivo y verificar si hay respuestas
    const formulario = FormApp.openByUrl(hoja.getRange(PARAMETROS.url).getValue());
    const respuestas = formulario.getResponses();
    
    if (respuestas.length == 0) throw 'No hay respuestas en el formulario.'

    // Señalizar inicio del proceso de extracción de respuestas
    hdc.toast('Obteniendo respuestas...', '', -1);

    // Posibles datos anteriores en gris claro durante el proceso
    let ultimaFila = hoja.getLastRow();
    if(ultimaFila > PARAMETROS.filaTabla) {
      hoja.getRange(PARAMETROS.filaTabla + 1,1, ultimaFila - PARAMETROS.filaTabla + 1, hoja.getLastColumn()).setFontColor('#d0d0d0');
      SpreadsheetApp.flush();
    }

    // Generar la fila de encabezado de la tabla de respuestas
    const encabezados = [];
    if (fechaSiNo) encabezados.push('🗓️ Marca tiempo');
    if (emailSiNo) encabezados.push('📨 Email');
    // Si numCampos > nº respuestas se toman todas las disponibles
    respuestas[0].getItemResponses().slice(0, numCampos).forEach(item => encabezados.push(item.getItem().getTitle()));
    if (urlSiNo) encabezados.push('✍️ URL edición');

    // Extraer respuestas
    const datos = respuestas.map(respuesta => {

      const filaDatos = [];
      if (fechaSiNo) filaDatos.push(respuesta.getTimestamp());
      if (emailSiNo) filaDatos.push(respuesta.getRespondentEmail());
      respuesta.getItemResponses().slice(0, numCampos).forEach(item => {
          // getResponse() pude devolver String | String | String[][], así que se aplana el array con profundidad 2,
          // como simple precaución, ver https://developers.google.com/apps-script/reference/forms/item-response#getresponse
        filaDatos.push(
          Array.isArray(item.getResponse())
        ? item.getResponse().flat().join(', ')
        : item.getResponse());
      });
      if (urlSiNo) filaDatos.push(respuesta.getEditResponseUrl());
      return filaDatos;
    
    });

    // Montar encabezado y respuestas en una sola tabla
    const tabla = [encabezados, ...datos];
       
    // Escribir tabla en la hoja de cálculo, borrando datos previos, si los hay
    if (hoja.getLastRow() > PARAMETROS.filaTabla) {
      hoja.getRange(PARAMETROS.filaTabla, 1, hoja.getLastRow() - PARAMETROS.filaTabla + 1, hoja.getLastColumn()).clearContent();
    }
    hoja.getRange(PARAMETROS.filaTabla, 1, tabla.length, tabla[0].length).setValues(tabla);

    // Informar del fin del proceso (con éxito)
    hdc.toast(`Respuestas obtenidas: ${tabla.length - 1}.`, '');
  
  } catch(e) {
    // Informar de error, si el objeto e es de tipo string es porque hemos llegado
    // aquí al fallar la comprobación de existencia de respuestas (¡sucio!).
    hdc.toast(typeof e == 'string' ? e: `Error interno: ${e.message}`, 'No hay respuestas en el formulario');
  
  } finally {
    // Esto se ejecuta siempre, tanto si hemos cazado algún error como si todo ha ido ok,
    // contenido de la tabla en color habitual.
    ultimaFila = hoja.getLastRow()
    if (ultimaFila > PARAMETROS.filaTabla) {
      hoja.getRange(PARAMETROS.filaTabla + 1,1, hoja.getLastRow() - PARAMETROS.filaTabla + 1, hoja.getLastColumn()).setFontColor(null);
    }
  }

}
