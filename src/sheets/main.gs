/**
 * @file    main.gs
 * @brief   Describes the routines to follow in order to retrieve the upcoming events' respective dates.
 * @author  Pedro Henrique Pinto de Oliveira
 * @date    2024-03-04
 */

/* Constants */
const ss = SpreadsheetApp.getActiveSpreadsheet();         // gets spreadsheet
const mainSheet = ss.getSheets()[0];                      // gets first sheet
const globalUpcomingDatesRange = "B3:B1000";              // all the possible cells which can be possibly filled with an event's date
const todayDay = new Date().getDate();
const sheet = ss.getSheets()[new Date().getMonth() + 1];  // getMonth() returns 0-11 months, not 1-12

/* Functions */
function onOpen(e)
{
  SpreadsheetApp.getUi().createMenu("Atualizar")
  .addItem("Atualizar agora", "main")
  .addToUi();
}

function findEvent_returnDate()
{
  let data = sheet.getDataRange().getValues();            // macro to retrieve (int row, int col) cell data
  let nDatesWritten = 0;

  for(let i = (tableHeader_rowValue_offset+todayDay); i <= 33; i++) {
    // Columns C:G

    for(let j = 3; j <= 7; j++) {
      if(sheet.getRange(i, j).getValue() == '') {} else {
        mainSheet.getRange(3 + nDatesWritten, 2).setValue(sheet.getRange(i, 1).getValue());
        nDatesWritten++;
      }
    }
  }
}

function clearOutdatedEvents()
{
  mainSheet.getRange(globalUpcomingDatesRange).clear();
}

function highlightToday_sEntry()
{
  let sheet = ss.getSheets()[new Date().getMonth() + 1];
  let dd = 2 + new Date().getDate();
  
  if(dd == (2 + 1)){
    let sheet_old = ss.getSheets()[new Date().getMonth() + 1 - 1];
    sheet_old.getRange("A31:A33").clearFormat();          // ranges from the 29th to the 31st day of the month
  } else {
    sheet.getRange(dd-1, 1).clearFormat();
  }
  
  sheet.getRange(dd, 1).setBackgroundRGB(255, 255, 0);
}

function main()
{
  console.log(setup(true));
  highlightToday_sEntry();
  clearOutdatedEvents();
  findEvent_returnDate();
}
