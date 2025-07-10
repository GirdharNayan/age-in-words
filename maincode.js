function ageInWord() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const inputColumn = 5;   
  const outputColumn = 6;  

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;  

  const numRows = lastRow - 1;
  const dateValues = sheet.getRange(2, inputColumn, numRows).getValues();

  const today = new Date();

  for (let i = 0; i < dateValues.length; i++) {
    let cellDate = dateValues[i][0];

    if (cellDate instanceof Date && !isNaN(cellDate)) {
      let elapsed = getDataString(cellDate, today);
      sheet.getRange(i + 2, outputColumn).setValue(elapsed); 
    } else {
      sheet.getRange(i + 2, outputColumn).clearContent();
    }
  }
}

function getDataString(startDate, endDate) {
  let years = endDate.getFullYear() - startDate.getFullYear();
  let months = endDate.getMonth() - startDate.getMonth();
  let days = endDate.getDate() - startDate.getDate();

  if (days < 0) {
    months--;
    let prevMonth = new Date(endDate.getFullYear(), endDate.getMonth(), 0);
    days += prevMonth.getDate();
  }
  if (months < 0) {
    years--;
    months += 12;
  }

  return years + " years, " + months + " months and " + days + " days";
}
