function myFunction() {
  spreadsheetId = '1oJjNkDS7-xIywtjYj-xR9EJvDRCmyYZeYc77ot7Zo_s';//Id del form
  readRange(spreadsheetId);
  //var calendars = Calendar.CalendarList.list();
  //for (var i = 0; i < calendars.items.length; i++) {
  //      Logger.log(calendars.items[i].summary);
  //      if(calendars.items[i].summary == "Administrativo"){
  //          var id = calendars.items[i].id;
  //          
  //      } 
  //    }
  // var events = Calendar.Events.list(id, {
  //  singleEvents: true,q: "lucasblanco3107@gmail.com"
  //});    
}

/**
 * Read a range (A1:D5) of data values. Logs the values.
 * @param {string} spreadsheetId The spreadsheet ID to read from.
 */
function readRange(spreadsheetId) {
  var response = Sheets.Spreadsheets.Values.get(spreadsheetId, 'Eventos');
  for (var i = 1; i < response.values.length; i++) {
    if(response.values[i][14] != response.values[i][15]){
      if(response.values[i][15] == "SI"){
        generateEvents(response.values[i]);
        Logger.log(response.values[i][15]);
      }
      
      
      
      if(response.values[i][15] == "NO"){Logger.log(response.values[i][15]);}
      if(response.values[i][15] == "QUITAR"){Logger.log(response.values[i][15]);}
      
      //var toStartDay = formatDay(response.values[i][8]); 
      //var startDay = new Date(toStartDay);
      //var toEndDay = formatDay(response.values[i][10]);
      //var endDay =  new Date(toEndDay);
      /*
      var startDay = formatDay(response.values[i][8]);
      var endDay = formatDay(response.values[i][10]);
      
      
      var days = getDays(response.values[i][12]);
      var tempDay = startDay;
      while(true){
        if(validateDay(days, new Date(tempDay).getDay())){
          var start = Utilities.formatDate(tempDay, 'America/Argentina/Cordoba', 'yyyy-MM-dd') + 'T'+ response.values[i][11] + '-03:00';
          var end = Utilities.formatDate(tempDay, 'America/Argentina/Cordoba', 'yyyy-MM-dd') + 'T'+ response.values[i][9] + '-03:00';
         
          createEvent(response.values[i][7], start , end , response.values[i][13], 
                      response.values[i][2] + " " + response.values[i][3] , response.values[i][6]);
          response.values[i][14] = response.values[i][15];

        }
        tempDay.setDate(tempDay.getDate() + 1 );
        if(tempDay > endDay){
          break;
        }
      }
      var qrpic = createQR(response.values[i][2] , response.values[i][3],  response.values[i][13] ,  response.values[i][4],  response.values[i][5] );

      sendMail(response.values[i][13], qrpic, response.values[i][1], response.values[i][2], response.values[i][3], response.values[i][4], response.values[i][5], response.values[i][6], response.values[i][7], response.values[i][8], response.values[i][10], response.values[i][11], response.values[i][9], response.values[i][12]);
   */ }
    
    
  }

  
  var valueRange = Sheets.newValueRange();
  valueRange.values = response.values;
  var result = Sheets.Spreadsheets.Values.update(valueRange, spreadsheetId, 'Eventos', {valueInputOption: "RAW"});
  Logger.log(response.values);
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

function sendMail(p_emailI, sintax, p_emailA, p_nombre, p_apellido, p_dni, p_telefono, p_espacio, p_evento, p_iHab, p_fHab, p_iEv, p_fEv, p_dias){

  //noReply	Boolean	true if the email should be sent from a generic no-reply email address to discourage recipients from responding to emails; this option is only possible for Google Workspace accounts, not Gmail users
    

//MAIL INVITADO
////////////////////////////////////
  var response = UrlFetchApp.fetch(sintax).getBlob().setName("vCard");

  MailApp.sendEmail({
    to: p_emailI,
    //name: "Acceso FCEFyN",
    subject: "Acceso FCEFyN - Invitación y vCard",
    htmlBody: "Hola, fuiste autorizado/a a ingresar en la FCEFyN. <br><br>" +
              "Autorización: " + p_evento +
              "<br>Inicio de Habilitación: " + p_iHab +
              "<br>Fin de Habilitación: " + p_fHab + 
              "<br>Inicio Horario de Habilitación: " + p_iEv +
              "<br>Fin Horario de Habilitación: " + p_fEv +
              "<br>Dias: " + p_dias +
              "<br>Con el código QR adjunto en el mail vas a poder ingresar.<br>" +  
              "<br>Gracias, Saludos!",
    inlineImages:
      {
        vCard: response
      }
  });
//////////////////////////////////////

//MAIL ADMINISTRATIVO
//////////////////////////////////////
  MailApp.sendEmail({
    to: p_emailA,
    //name: "Acceso FCEFyN",
    subject: "Acceso FCEFyN - Creación de Autorización",
    htmlBody: "Hola, acabas de autorizar un ingreso en la FCEFyN. <br><br>" +
              "Espacio: " + p_espacio +   
              "<br>Nombre Autorización: " + p_evento +
              "<br><br>Datos del Invitado:" + 
              "<br>Nombre: " + p_nombre +
              "<br>Apellido: " + p_apellido +
              "<br>DNI: " + p_dni +
              "<br>Mail: " + p_emailI +
              "<br>Telefono: " + p_telefono +
              "<br><br>Datos de la Autorización:" + 
              "<br>Inicio de Habilitación: " + p_iHab +
              "<br>Fin de Habilitación: " + p_fHab + 
              "<br>Inicio Horario de Habilitación: " + p_iEv +
              "<br>Fin Horario de Habilitación: " + p_fEv +
              "<br>Dias: " + p_dias +
              "<br><br>Gracias, Saludos! <br>",
  });

//////////////////////////////////////
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

      var startDay = formatDay(response.values[8]);
      var endDay = formatDay(response.values[10]);
      var days = getDays(response.values[12]);
      var tempDay = startDay;

      while(true){
        if(validateDay(days, new Date(tempDay).getDay())){
          var start = Utilities.formatDate(tempDay, 'America/Argentina/Cordoba', 'yyyy-MM-dd') + 'T'+ response.values[11] + '-03:00';
          var end = Utilities.formatDate(tempDay, 'America/Argentina/Cordoba', 'yyyy-MM-dd') + 'T'+ response.values[9] + '-03:00';
         
          createEvent(response.values[7], start , end , response.values[13], 
                      response.values[2] + " " + response.values[3] , response.values[6]);
          response.values[14] = response.values[15];

        }
        tempDay.setDate(tempDay.getDate() + 1 );
        if(tempDay > endDay){
          break;
        }
      }
}
 
