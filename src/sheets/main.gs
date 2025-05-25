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

let nDatesWritten = 0;

/* Functions */
function onOpen(e)
{
  SpreadsheetApp.getUi().createMenu("Atualizar")
  .addItem("Atualizar agora", "main")
  .addToUi();
}

function writeEventsWithinInterval(sourceSheet, beginIndex, endIndex)
{
  for(let i = beginIndex; i <= endIndex; i++)
  {
    // Columns C:G (5 possible event entries)
    for(let j = 3; j <= 7; j++)
    {
      if(sourceSheet.getRange(i, j).getValue() != '')
      {
        mainSheet.getRange(tableHeader_rowValue_offset + (++nDatesWritten), 2).setValue(sourceSheet.getRange(i, 1).getValue());
      }
    }
  }
}

function findEvent_returnDate()
{
  let data = sheet.getDataRange().getValues();            // macro to retrieve (int row, int col) cell data
  let nextSheet = ss.getSheets()[new Date().getMonth() + 1 + 1];
  let nDatesWritten = 0;

  writeEventsWithinInterval(sheet, tableHeader_rowValue_offset + todayDay, tableHeader_rowValue_offset + 31);

  if(sheet.getRange(tableHeader_rowValue_offset+todayDay+8, 1).isBlank()) // if no data is found after the next seven days, begin showing next month's schedule
  {
    writeEventsWithinInterval(nextSheet, tableHeader_rowValue_offset + 1, tableHeader_rowValue_offset + 1 + 7);
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
    sheet.getRange(dd-1, 1).setFontSize(11); // New
  }

  sheet.getRange(dd, 1).setFontSize(11); // New
  sheet.getRange(dd, 1).setBackgroundRGB(255, 255, 0);
}

function main()
{
  console.log(setup(true));
  highlightToday_sEntry();
  clearOutdatedEvents();
  findEvent_returnDate();
}

