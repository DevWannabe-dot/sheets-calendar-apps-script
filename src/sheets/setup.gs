/**
 * @file    setup.gs
 * @brief   This run once script creates sample events for the next three days to help users understand the functionality.
 * @author  Pedro Henrique Pinto de Oliveira
 * @date    2024-03-16
 */

/* Constants */
// Global scope constants borrowed from main.gs
const entry_firstEventColumn = 3;
const tableHeader_rowValue_offset = 2;

/* Functions */
function is_LeapYear(year)
{
  if(year % 400 === 0 || (year % 4 === 0 && year % 100 !== 0))
  {
    return true;
  }
  return false;
}

function find_AmountOfDaysLeftInMonth(date, month, year)
{
  switch (month){
    case 1:
    case 3:
    case 5:
    case 7:
    case 8:
    case 10:
    case 12:
      return (31 - date);
    break;
    case 2:
      if(is_LeapYear(year)){
        return (29 - date);
      }
      return (28 - date);
    break;
    case 4:
    case 6:
    case 9:
    case 11:
      return (30 - date);
    break;
    default:
    break;
  }
}

function setup(mainFunctionCall) {
  // If sheet already contains values, must NOT setup
  if(!(mainSheet.getRange(3, 2).isBlank())) return false;

  const sampleEvents = ["Evento 1", "Evento 2", "Evento 3"];

  let year = new Date().getFullYear();
  let month = new Date().getMonth() + 1;                                      // returns 0-11 months
  let nextMonth = month + 1;
  let date = new Date().getDate();
  let amount = find_AmountOfDaysLeftInMonth(date, month, year);

  for(let i = date + tableHeader_rowValue_offset; i < (date+amount); i++){
    if(sampleEvents.length === 0) break;
    sheet.getRange(i, entry_firstEventColumn).setValue(sampleEvents.shift()); // allows stack-based JavaScript array to behave like a queue
  }
  if(amount < 3 && sampleEvents.length === 0){
    let i = 1 + tableHeader_rowValue_offset;                                  // Begins at '@NextMonth'!C3's address
    while(sampleEvents.length() > 0){
      ss.getSheets()[nextMonth].getRange(i, entry_firstEventColumn).setValue(sampleEvents.shift());
    }
  }

  if(!mainFunctionCall) main();                                               // If main() didn't call setup(), then main() must be run to findEvent_returnDate()s (might have been onOpen(e) trigger)

  return true;
}
