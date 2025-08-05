function syncEventsFromSheets() {
  var lock = LockService.getScriptLock();
  try {
    // Wait up to 30 seconds for the lock to prevent concurrent runs.
    lock.waitLock(30000);

    // Mapping from person name to their specific Calendar ID.
    // Please update these IDs with your actual calendar IDs these IDs are mine.
    var calendarMapping = {
      "John": "Your Sub Calendar ID @group.calendar.google.com",
      "Brian": "Your Sub Calendar ID @group.calendar.google.com",
      "Mateo": "Your Sub Calendar ID @group.calendar.google.com",
      "Dorn": "Your Sub Calendar ID @group.calendar.google.com"
      // Add additional mappings as needed.
    };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();

    // Map some common hex background colors to CalendarApp event colors.
    var colorMap = {
      "#ffff00": CalendarApp.EventColor.YELLOW,  // yellow
      "#ff0000": CalendarApp.EventColor.RED,       // red
      "#00ff00": CalendarApp.EventColor.GREEN,     // green
      "#0000ff": CalendarApp.EventColor.BLUE,      // blue
      "#ffa500": CalendarApp.EventColor.ORANGE,     // orange
      "#800080": CalendarApp.EventColor.PURPLE      // purple
    };

    // Process each sheet that follows the naming pattern "Month Year"
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      var sheetName = sheet.getName();
      var parts = sheetName.split(" ");
      if (parts.length < 2) {
        Logger.log("Skipping sheet: " + sheetName);
        continue;
      }

      var monthName = parts[0];
      var year = parseInt(parts[1], 10);
      if (isNaN(year)) {
        Logger.log("Invalid year in sheet: " + sheetName);
        continue;
      }

      // Lookup table for number of days in each month (handles February leap-year check)
      var daysInMonthLookup = {
        "January": 31,
        "February": ((year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0)) ? 29 : 28),
        "March": 31,
        "April": 30,
        "May": 31,
        "June": 30,
        "July": 31,
        "August": 31,
        "September": 30,
        "October": 31,
        "November": 30,
        "December": 31
      };
      var daysInMonth = daysInMonthLookup[monthName];
      if (!daysInMonth) {
        Logger.log("Invalid month in sheet: " + sheetName);
        continue;
      }

      Logger.log("Processing sheet: " + sheetName + " (" + monthName + " " + year + "), days in month: " + daysInMonth);

      // Get day headers from row 4 starting at column C (day numbers)
      var lastCol = sheet.getLastColumn();
      var dayHeaders = sheet.getRange(4, 3, 1, lastCol - 2).getDisplayValues()[0];

      // Get data rows: names in column B (starting at row 6) and event cells in columns C onward
      var lastRow = sheet.getLastRow();
      var dataRange = sheet.getRange(6, 2, lastRow - 5, lastCol - 1);
      var data = dataRange.getDisplayValues();
      // Also get background colors from column B for each row.
      var bgColors = sheet.getRange(6, 2, lastRow - 5, 1).getBackgrounds();

      // Define the date range for this sheet (covers the entire month)
      var monthStart = new Date(year, getMonthIndex(monthName), 1);
      var monthEnd = new Date(year, getMonthIndex(monthName) + 1, 1);

      // Process each person (each row)
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var personName = row[0].trim();
        if (!personName) continue;

        // Look up the Calendar ID for this person.
        var calendarId = calendarMapping[personName];
        if (!calendarId) {
          Logger.log("No calendar mapping found for: " + personName);
          continue;
        }
        var calendar = CalendarApp.getCalendarById(calendarId);
        if (!calendar) {
          Logger.log("Could not get calendar with ID for: " + personName);
          continue;
        }

        // Determine event color based on the background color in column B.
        var bg = bgColors[i][0].toLowerCase();
        var eventColor = colorMap[bg] || null;

        // Build a collection of desired events for this person.
        var desiredEvents = {};
        var tz = Session.getScriptTimeZone();

        // Process event cells (starting with Column C which is index 1 in the data row)
        var j = 1;
        while (j < row.length) {
          var cellVal = row[j] ? row[j].trim() : "";
          if (cellVal === "H") {
            // Group consecutive "H" cells for a full-day event.
            var startCol = j;
            var endCol = j;
            while (endCol + 1 < row.length && row[endCol + 1] && row[endCol + 1].trim() === "H") {
              endCol++;
            }
            var startDayStr = dayHeaders[startCol - 1].trim();
            var endDayStr = dayHeaders[endCol - 1].trim();
            var startDay = parseInt(startDayStr, 10);
            var endDay = parseInt(endDayStr, 10);

            // Validate day values against daysInMonth.
            if (isNaN(startDay) || isNaN(endDay) || startDay > daysInMonth || endDay > daysInMonth) {
              Logger.log("Invalid day header in sheet " + sheetName + " at columns " + (startCol + 2) + " to " + (endCol + 2));
              j = endCol + 1;
              continue;
            }

            var eventStart = new Date(year, getMonthIndex(monthName), startDay);
            // For all-day events, the end date is exclusive; add one day to the final day.
            var eventEnd = new Date(year, getMonthIndex(monthName), endDay + 1);
            var formattedStart = Utilities.formatDate(eventStart, tz, "yyyy-MM-dd");
            var formattedEnd = Utilities.formatDate(eventEnd, tz, "yyyy-MM-dd");
            // Unique key includes sheet name, row number, event type, and formatted dates.
            var key = sheetName + "|" + (i + 6) + "|H|" + formattedStart + "|" + formattedEnd;

            desiredEvents[key] = {
              key: key,
              title: personName + " - OOO",
              start: eventStart,
              end: eventEnd,
              allDay: true,
              color: eventColor,
              type: "H"
            };
            j = endCol + 1;
          } else if (cellVal === "H1" || cellVal === "H2") {
            // Process half-day events individually.
            var dayStr = dayHeaders[j - 1].trim();
            var dayNum = parseInt(dayStr, 10);
            if (isNaN(dayNum) || dayNum > daysInMonth) {
              Logger.log("Invalid day value '" + dayStr + "' at column " + (j + 2) + " in sheet " + sheetName);
              j++;
              continue;
            }
            var eventStart, eventEnd;
            if (cellVal === "H1") {
              // Morning half-day (adjust times as needed)
              eventStart = new Date(year, getMonthIndex(monthName), dayNum, 8, 0);
              eventEnd = new Date(year, getMonthIndex(monthName), dayNum, 12, 0);
            } else {
              // Afternoon half-day
              eventStart = new Date(year, getMonthIndex(monthName), dayNum, 13, 0);
              eventEnd = new Date(year, getMonthIndex(monthName), dayNum, 17, 0);
            }
            var formattedEventDate = Utilities.formatDate(eventStart, tz, "yyyy-MM-dd");
            var key = sheetName + "|" + (i + 6) + "|Half|" + formattedEventDate + "|" + cellVal;
            desiredEvents[key] = {
              key: key,
              title: personName + " - Half",
              start: eventStart,
              end: eventEnd,
              allDay: false,
              color: eventColor,
              type: "Half"
            };
            j++;
          } else {
            j++;
          }
        }

        // Retrieve existing events in the target calendar for this sheet’s month that were created by our sync.
        var existingEvents = calendar.getEvents(monthStart, monthEnd);
        var existingMap = {};
        for (var e = 0; e < existingEvents.length; e++) {
          var ev = existingEvents[e];
          var desc = ev.getDescription();
          if (desc && desc.indexOf("SheetSync: ") === 0) {
            var evKey = desc.substring("SheetSync: ".length).trim();
            existingMap[evKey] = ev;
          }
        }

        // Create or update events based on desiredEvents.
        for (var key in desiredEvents) {
          var evData = desiredEvents[key];
          if (existingMap.hasOwnProperty(key)) {
            // Update existing event if needed.
            var ev = existingMap[key];
            if (ev.getTitle() !== evData.title) {
              ev.setTitle(evData.title);
            }
            if (evData.allDay) {
              ev.setAllDayDates(evData.start, evData.end);
            } else {
              ev.setTime(evData.start, evData.end);
            }
            if (evData.color && ev.getColor() !== evData.color) {
              ev.setColor(evData.color);
            }
            // Remove from existingMap so that remaining events can be deleted.
            delete existingMap[key];
          } else {
            // Create a new event.
            var newEvent;
            if (evData.allDay) {
              newEvent = calendar.createAllDayEvent(evData.title, evData.start, evData.end);
            } else {
              newEvent = calendar.createEvent(evData.title, evData.start, evData.end);
            }
            if (evData.color) {
              newEvent.setColor(evData.color);
            }
            // Tag the event with our unique key in its description.
            newEvent.setDescription("SheetSync: " + key);
          }
        }

        // Delete any existing events that are no longer desired.
        for (var key in existingMap) {
          try {
            existingMap[key].deleteEvent();
          } catch (ex) {
            Logger.log("Error deleting event with key " + key + ": " + ex);
          }
        }
      } // end for each person
    } // end for each sheet
  } catch (err) {
    Logger.log("Error in syncEventsFromSheets: " + err);
  } finally {
    lock.releaseLock();
  }
}

// Helper function: Convert full month name to month index (0-based)
function getMonthIndex(monthName) {
  var monthMap = {
    "January": 0,
    "February": 1,
    "March": 2,
    "April": 3,
    "May": 4,
    "June": 5,
    "July": 6,
    "August": 7,
    "September": 8,
    "October": 9,
    "November": 10,
    "December": 11
  };
  return monthMap[monthName];
}
