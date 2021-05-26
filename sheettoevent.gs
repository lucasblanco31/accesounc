function myFunction() {
  spreadsheetId = '1oJjNkDS7-xIywtjYj-xR9EJvDRCmyYZeYc77ot7Zo_s';//Id del form
  readRange(spreadsheetId);
}

/**
 * Read a range (A1:D5) of data values. Logs the values.
 * @param {string} spreadsheetId The spreadsheet ID to read from.
 */
function readRange(spreadsheetId) {
  var response = Sheets.Spreadsheets.Values.get(spreadsheetId, 'Eventos');
  for (var i = 1; i < response.values.length; i++)
  {
    if(response.values[i][14] == response.values[i][15])
    {
      //Respuesta sin revisar o autorizacion sin cambios
      Logger.log("Por revisar: " + i + "," + response.values[14] + "," + response.values[15]);
      continue;
    }
    if(response.values[i][15] == "SI"){
        generateEvents(response.values[i]);
        var qrpic = createQR(response.values[i][2] , response.values[i][3],  response.values[i][13] ,  response.values[i][4],       response.values[i][5] );
        sendMail2(response.values[i], "SI", qrpic);
        response.values[i][14] = response.values[i][15];
        Logger.log(response.values[i][15]);
    }
    if(response.values[i][15] == "NO"){
      if(response.values[i][14] == "SI"){
        quitarEvento(response.values[i]);
        sendMail2(response.values[i], "BAJA", null);
        response.values[i][14] = response.values[i][15];
      }
      else{
        sendMail2(response.values[i], "NO", null);
        response.values[i][14] = response.values[i][15];
      }
     
      Logger.log(response.values[i][15]);
    }
  } 
  var valueRange = Sheets.newValueRange();
  valueRange.values = response.values;
  var result = Sheets.Spreadsheets.Values.update(valueRange, spreadsheetId, 'Eventos', {valueInputOption: "RAW"});
}

//calendarId, start, end, email, etc...
function createEvent(p_summary , p_start, p_end , p_email, p_displayname, p_espacio) {
  
  var calendarId = getCalendarID(p_espacio);
  var event = {
    summary: p_summary,
    start: {
      dateTime: p_start 
    },
    end: {
      dateTime: p_end 
    },
    attendees: [
      {email: p_email, displayName: p_displayname }
    ],
    // Red background. Use Calendar.Colors.get() for the full list.
  };
  event = Calendar.Events.insert(event, calendarId);
  Logger.log('Event ID: ' + event.id);
}
function getDays(days){
  var day_array = [{}];
  var day_num_array = [{}];
  day_array = days.split(", ");
  for(var i = 0 ; i < day_array.length; i++){
    switch (day_array[i]){
      case "Lunes":
          day_num_array[i] = 1;
          break;
      case "Martes":
          day_num_array[i] = 2;
          break;
      case "Miercoles":
          day_num_array[i] = 3;
          break;
      case "Jueves":
          day_num_array[i] = 4;
          break;
      case "Viernes":
          day_num_array[i] = 5;
          break;
      case "Sabado":
          day_num_array[i] = 6;
          break;
    }

  }
  return day_num_array
}
function validateDay(val_days , day){
   for(var i = 0; i < val_days.length ; i++){
     if(val_days[i] == day){
       return true;
     }
   }
   return false;
}
function createQR(p_nombre, p_apellido, p_email, p_dni, p_telefono){

    var msg = 'BEGIN:VCARD\nVERSION:3.0\r\nN;' + p_nombre + ';' + p_apellido + '\r\nFN;' + p_nombre + ' ' +         p_apellido + '\r\nEMAIL;CHARSET=UTF-8;type=HOME;INTERNET:' + p_email + '\r\nNOTE;CHARSET=UTF-8:ARG_DNI=' + p_dni + '\r\nTEL;TYPE=voice,work,pref:' + p_telefono + '\nEND:VCARD';

    var sintax = 'https://chart.googleapis.com/chart?cht=qr&chs=512x512&chl=' + urlEncode(msg);
    Logger.log(urlEncode(msg))
    Logger.log(sintax);
    return sintax;
    //sendMail(p_email , sintax);
}

function urlEncode(str) {
  var forAsciiConvert =
    "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz*-.@_";
  var conv = Utilities.newBlob(forAsciiConvert)
    .getBytes()
    .map(function(e) {
      return ("0" + (e & 0xff).toString(16)).slice(-2);
    });
  var bytes = Utilities.newBlob("")
    .setDataFromString(str, "Shift_JIS")
    .getBytes();
  return (res = bytes
    .map(function(byte) {
      var n = ("0" + (byte & 0xff).toString(16)).slice(-2);
      return conv.indexOf(n) != -1
        ? String.fromCharCode(
            parseInt(n[0], 16).toString(2).length == 4
              ? parseInt(n, 16) - 256
              : parseInt(n, 16)
          )
        : ("%" + n).toUpperCase();
    })
    .join(""));
}
function sendMail2(response, option, sintax){
    if(option == "SI"){
      var qrimage = UrlFetchApp.fetch(sintax).getBlob().setName("vCard");
        MailApp.sendEmail({
          //noReply: true,
          to: response[13],
          name: "Acceso FCEFyN",
          subject: "Acceso FCEFyN - Invitación y vCard",
          htmlBody: "Hola, fuiste autorizado/a a ingresar en la FCEFyN. <br><br>" +
              "Nombre Autorización: " + response[7] +
              "<br>Inicio de Habilitación: " + response[8] +
              "<br>Fin de Habilitación: " + response[10] + 
              "<br>Inicio Horario de Habilitación: " + response[11] +
              "<br>Fin Horario de Habilitación: " + response[9] +
              "<br>Dias: " + response[12] +
              "<br>Guardá éste código QR adjunto en el mail para presentar en la entrada de la facultad.<br>" +  
              "<br>Gracias, Saludos!",
          inlineImages:
          {
            vCard: qrimage
          }});
    }
    if(option == "NO"){
      //Se rechazo la autorizacion
        MailApp.sendEmail({
          //noReply: true,
          to: response[13],
          name: "Acceso FCEFyN",
          subject: "Acceso FCEFyN - Rechazo de Autorización",
          htmlBody: "Hola, tu pedido de autorización para ingresar a la FCEFyN fue rechazado. <br><br>" +
              "Nombre Autorización: " + response[7] +
              "<br>Inicio de Habilitación: " + response[8] +
              "<br>Fin de Habilitación: " + response[10] + 
              "<br>Inicio Horario de Habilitación: " + response[11] +
              "<br>Fin Horario de Habilitación: " + response[9] +
              "<br>Dias: " + response[12] +
              "<br>Comunicate con la FCEFyN para saber más sobre tu rechazo.<br>" +  
              "<br>Gracias, Saludos!",   
          });
    }
    if(option == "BAJA"){
      //Se quito la autorizacion
        MailApp.sendEmail({
          //noReply: true,
          to: response[13],
          name: "Acceso FCEFyN",
          subject: "Acceso FCEFyN - Autorización dada de baja",
          htmlBody: "Hola, tu autorización para ingresar en la FCEFyN fue dada de baja. <br><br>" +
              "Nombre Autorización: " + response[7] +
              "<br>Inicio de Habilitación: " + response[8] +
              "<br>Fin de Habilitación: " + response[10] + 
              "<br>Inicio Horario de Habilitación: " + response[11] +
              "<br>Fin Horario de Habilitación: " + response[9] +
              "<br>Dias: " + response[12] +
              "<br>Comunicate con la FCEFyN para saber porqué fue dada de baja.<br>" +
              "<br>Gracias, Saludos!",
          });
    }
}
function getCalendarID(p_espacio) {
  var calendars;
  var pageToken;
  do {
    calendars = Calendar.CalendarList.list({
      maxResults: 100,
      pageToken: pageToken
    });
    if (calendars.items && calendars.items.length > 0) {
      for (var i = 0; i < calendars.items.length; i++) {
        var calendar = calendars.items[i];
        Logger.log('%s (ID: %s)', calendar.summary, calendar.id);
        if (calendar.summary == p_espacio){
          return calendar.id;
        }
      }
    } else {
      Logger.log('No calendars found.');
    }
    pageToken = calendars.nextPageToken;
  } while (pageToken);
}
function formatDay(fecha){
  var meses = Array(12);
  meses = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sept', 'oct', 'nov', 'dic'];  
  var vfecha = fecha.split("/");

  var rfecha = meses[vfecha[1]-1];
  rfecha = rfecha + '/' + vfecha[0] + '/' + vfecha[2];

  
  return new Date(rfecha);
}
function generateEvents(response){

      var startDay = formatDay(response[8]);
      var endDay = formatDay(response[10]);
      var days = getDays(response[12]);
      var tempDay = startDay;

      while(true){
        if(validateDay(days, new Date(tempDay).getDay())){
          var start = Utilities.formatDate(tempDay, 'America/Argentina/Cordoba', 'yyyy-MM-dd') + 'T'+ response[11] + '-03:00';
          var end = Utilities.formatDate(tempDay, 'America/Argentina/Cordoba', 'yyyy-MM-dd') + 'T'+ response[9] + '-03:00';
         
          createEvent(response[7], start , end , response[13], 
                      response[2] + " " + response[3] , response[6]);
          //response.values[14] = response.values[15];

        }
        tempDay.setDate(tempDay.getDate() + 1 );
        if(tempDay > endDay){
          break;
        }
      }
}
function quitarEvento(response){
  
  var date = Utilities.formatDate(new Date(), 'America/Argentina/Cordoba', 'yyyy-MM-dd') + 'T00:00:00-03:00';
  
  var calendars = Calendar.CalendarList.list();
  for (var i = 0; i < calendars.items.length; i++){
        Logger.log(calendars.items[i].summary);
        if(calendars.items[i].summary == response[6]){
            var calendarID = calendars.items[i].id;
            break;
        } 
  }
  var events = Calendar.Events.list(calendarID, {
    timeMin: date,
    singleEvents: true,
    q: response[13],
  });

  var eventsID = new Array();
  var indice = 0;
  for(var j=0; j<events.items.length; j++){
    if(events.items[j].summary == response[7]){
      eventsID[indice] = events.items[j].id;
      Logger.log("calid:" + eventsID[indice] + ", name:" + events.items[j].summary);
      indice++;
    }    
  }
  for(var k=0; k<eventsID.length; k++){
    Calendar.Events.remove(calendarID, eventsID[k]);      
  }
  Logger.log(events);    
}
