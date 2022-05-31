/********************************IMPORTANTE*************************************/
// Antes de utilizar este script, es importante habilitar el API de Youtube,
// Menú "Resources" > Advanced Google Services y Habilitar YouTube Data API v3
// Yotube API Documentation: https://developers.google.com/youtube/v3/docs/videos/list
/********************************************************************************/

/* Esta función agrega un elmento al menú llamado "Youtube Report"
Docuentación sobre elementos de menú: https://developers.google.com/apps-script/quickstart/custom-functions
*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Reporte Youtube 🔴')
      .addItem('Actualizar reporte 📊', 'menuItem1')
      .addToUi();
}
function menuItem1() {
  updateStats();
  SpreadsheetApp.getUi().alert('El reporte ha sido actualizado');
}


/* Aquí inicia el código que hace el Query al Youtube API
Documento original: https://dev.to/rick_viscomi/using-sheets-and-the-youtube-api-to-track-video-analytics-6el
*/

// This is "Sheet1" by default. Keep it in sync after any renames.
var SHEET_NAME = 'Video Stats';

// This is the named range containing all video IDs.
var VIDEO_ID_RANGE_NAME = 'IDs';

// Update these values after adding/removing columns.
// Se agregaron los objetos PDATE y TITLE que se utilizan para mostrar la fecha de publicación y el título del video (respectivamente)

var Column = {
  PDATE: 'C',
  TITLE: 'D',
  VIEWS: 'E',
  LIKES: 'F',
  DISLIKES: 'G',
  COMMENTS: 'H',
  DURATION: 'I'
};

function updateStats() {
  var spreadsheet = SpreadsheetApp.getActive();
  var videoIds = getVideoIds();
  var stats = getStats(videoIds.join(','));
  writeStats(stats);
}

// Gets all video IDs from the range and ignores empty values.
function getVideoIds() {
  var spreadsheet = SpreadsheetApp.getActive();
  var range = spreadsheet.getRangeByName(VIDEO_ID_RANGE_NAME);
  var values = range.getValues();
  var videoIds = [];
  for (var i = 0; i < values.length; i++) {
    var value = values[i][0];
    if (!value) {
      return videoIds;
    }
    videoIds.push(value);
  }
  return videoIds;
}

// Queries the YouTube API to get stats for all videos.
function getStats(videoIds) {
  return YouTube.Videos.list('statistics,contentDetails,snippet', {'id': videoIds}).items;
}

// Converts the API results to cells in the sheet.
function writeStats(stats) {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  var durationPattern = new RegExp(/PT((\d+)M)?(\d+)S/);
  var datePattern = new RegExp(/PT((\d+)M)?(\d+)S/);
  for (var i = 0; i < stats.length; i++) {
    var cell = sheet.setActiveCell(Column.VIEWS + (2+i));
    cell.setValue(stats[i].statistics.viewCount);
    cell = sheet.setActiveCell(Column.LIKES + (2+i));
    cell.setValue(stats[i].statistics.likeCount);
    cell = sheet.setActiveCell(Column.DISLIKES + (2+i));
    cell.setValue(stats[i].statistics.dislikeCount);
    cell = sheet.setActiveCell(Column.COMMENTS + (2+i));
    cell.setValue(stats[i].statistics.commentCount);
    cell = sheet.setActiveCell(Column.DURATION + (2+i));
    var duration = stats[i].contentDetails.duration;
    var result = durationPattern.exec(duration);
    var min = result && result[2] || '00';
    var sec = result && result[3] || '00';
    cell.setValue('00:' + min + ':' + sec);
    //Estas útimas líneas no estan en el código original y fueron agregadas por @supirole para extraer el Título del video y la fecha de publicación.
    cell = sheet.setActiveCell(Column.TITLE + (2+i));
    cell.setValue(stats[i].snippet.title);
    cell = sheet.setActiveCell(Column.PDATE + (2+i));
    cell.setValue(stats[i].snippet.publishedAt.substring(0,10));
  }
}


// Este reporte tiene adicional al botón del menú customizado un trigger que se ejecuta cada vez que se abre el documento
// se pueden ver ó editar los triggers en el menú Edit > Current project's triggers
