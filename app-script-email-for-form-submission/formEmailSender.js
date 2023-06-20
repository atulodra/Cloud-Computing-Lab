function sendEmail(e) {
  let values = e.values
  const emailRecipient = values[1];
  const name = values[2];
  const body = `Dear ${name}, 
  Thank you for submitting your response to the google form. Your response is greatly appreciated!`;

  //Send email using GmailApp class
  GmailApp.sendEmail(emailRecipient, 'Google Form Subimission Notice', body);

}
