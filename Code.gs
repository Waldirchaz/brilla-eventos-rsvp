function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  template.modo = (e && e.parameter && e.parameter.modo) ? e.parameter.modo : 'invitado';
  template.grupoId = (e && e.parameter && e.parameter.grupo) ? e.parameter.grupo : '';
  
  return template.evaluate()
      .setTitle('Brilla Eventos - Sistema RSVP')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function obtenerDatosGrupo(grupoId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitados');
  var data = sheet.getDataRange().getValues();
  var miembros = [];
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === String(grupoId).trim()) {
      miembros.push({
        row: i + 1,
        id: data[i][0],
        nombre: data[i][1],
        grupoId: data[i][2],
        tipo: data[i][3],
        estado: data[i][4] || 'Pendiente'
      });
    }
  }
  return miembros;
}

function guardarRSVP(respuestas) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitados');
  respuestas.forEach(function(item) {
    sheet.getRange(item.row, 5).setValue(item.estado);
  });
  return { success: true, message: "¡Tus respuestas han sido guardadas con éxito!" };
}

function obtenerTodosLosInvitados() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitados');
  var data = sheet.getDataRange().getValues();
  var lista = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][1]) {
      lista.push({
        row: i + 1,
        id: data[i][0],
        nombre: data[i][1],
        grupoId: data[i][2],
        tipo: data[i][3],
        estado: data[i][4] || 'Pendiente',
        acompanantes: data[i][5] || 0,
        contacto: data[i][6] || ''
      });
    }
  }
  return lista;
}

function obtenerSiguienteGrupoId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitados');
  var data = sheet.getDataRange().getValues();
  var maxNum = 0;
  
  for (var i = 1; i < data.length; i++) {
    var gId = String(data[i][2]).trim();
    if (gId.indexOf('GRP-') === 0) {
      var num = parseInt(gId.replace('GRP-', ''), 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }
  
  var siguienteNum = maxNum + 1;
  return 'GRP-' + (siguienteNum < 10 ? '0' + siguienteNum : siguienteNum);
}

function agregarInvitado(datos) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitados');
  var grupoId = datos.grupoId ? datos.grupoId.trim() : '';
  
  if (!grupoId || grupoId.includes('Cargando')) {
    grupoId = obtenerSiguienteGrupoId();
  }
  
  datos.invitados.forEach(function(inv, index) {
    var lastRow = sheet.getLastRow() + 1;
    sheet.appendRow([
      lastRow,
      inv.nombre,
      grupoId,
      inv.tipo || (index === 0 ? 'Principal' : 'Vinculado'),
      'Pendiente',
      0,
      ''
    ]);
  });
  
  return { success: true, message: "Grupo de invitados agregado correctamente" };
}

function eliminarInvitado(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitados');
  sheet.deleteRow(row);
  return { success: true, message: "Invitado eliminado correctamente" };
}

function obtenerUrlWebApps() {
  return ScriptApp.getService().getUrl();
}
