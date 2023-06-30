function sendEmail() {
   let excel = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = excel.getSheetByName("Sheet1");
  // Getting last row
  let lastRow  = sheet.getLastRow();
  // Getting second last Column
  let secondLastColumn = sheet.getLastColumn()-1
  let range = sheet.getRange(2,2,lastRow-1,secondLastColumn)
  let values = range.getValues();

  //Get current month
  const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const d = new Date();
  const month = months[d.getMonth()];

  for (let i = 0; i < values.length; i++) {
    let status;
    let [name, email, salary] = values[i];
    if (Email) {
      try {
        let msg = buildBody(name, salary, month);
        MailApp.sendEmail(email, `Salary for month of ${month}`, msg);
        status = "success";
      } catch (err) {
        console.log(err);
        status = "Fail";
      }
    } else {
      status = "No Email";
    }
    let cell = range.getCell(i + 1, 4);
    cell.setValue(status);
  }
}

const buildBody = (name, salary, month) => {
  return `Hi ${name}, your salary for the month of ${month} has been credited. Salary:${salary}`;
};
