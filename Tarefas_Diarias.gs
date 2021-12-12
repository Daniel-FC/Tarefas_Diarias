let app = SpreadsheetApp;
let calendar = CalendarApp.getCalendarsByName("Diário")[0];
let sheet = app.getActiveSheet();

//Calendar
function myTasks() {
  let startSheet = "";
  let endSheet = "";

  if(new Date().getDay() == 0) {
    startSheet = "A3";
    endSheet = "C";
  } else if (new Date().getDay() == 1) {
    startSheet = "D3";
    endSheet = "F";
  } else if (new Date().getDay() == 2) {
    startSheet = "G3";
    endSheet = "I";
  } else if (new Date().getDay() == 3) {
    startSheet = "J3";
    endSheet = "L";
  } else if (new Date().getDay() == 4) {
    startSheet = "M3";
    endSheet = "O";
  } else if (new Date().getDay() == 5) {
    startSheet = "P3";
    endSheet = "R";
  } else if (new Date().getDay() == 6) {
    startSheet = "S3";
    endSheet = "U";
  }

  let yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1).toString;
  let myEvents = calendar.getEventsForDay(yesterday);
  myEvents.map(function(elem, ind, obj) {
    elem.deleteEvent();
  });

  let spaceRange = startSheet + ":" + endSheet;
  let range = sheet.getRange(spaceRange).getValues();
  range.map(function(elem, ind, obj) {
    if(elem[0] != "") {
      let dateStart = new Date();
      let timeStart = elem[1].split(":");

      let dateEnd = new Date();
      let timeEnd = elem[2].split(":");

      dateStart.setHours(timeStart[0]);
      dateStart.setMinutes(timeStart[1]);
      dateStart.setSeconds(timeStart[2]);

      dateEnd.setHours(timeEnd[0]);
      dateEnd.setMinutes(timeEnd[1]);
      dateEnd.setSeconds(timeEnd[2]);

      calendar.createEvent(elem[0], dateStart, dateEnd);
    }
  });
}
