function dropdowns() {
  var IndiapostDD = SpreadsheetApp.newDataValidation()
    .setHelpText('Click and enter a value from range')
    .requireValueInRange(IndiapostSheet.getRange('India Post!$A$4:$A'), true)
    .setAllowInvalid(false)
    .build();

  var ShiprocketDD = SpreadsheetApp.newDataValidation()
    .setHelpText('Click and enter a value from range')
    .requireValueInRange(ShiprocketSheet.getRange('Shiprocket (with Gst)!$B$1:$1'), true)
    .setAllowInvalid(false)
    .build();

  var MailInnoDD = SpreadsheetApp.newDataValidation()
    .setHelpText('Click and enter a value from range')
    .requireValueInRange(MailInnoSheet.getRange('Mail Innovation!$B$1:$1'), true)
    .setAllowInvalid(false)
    .build();

  var BlankVal = SpreadsheetApp.newDataValidation()
    .setHelpText('Click and enter a value from range')
    .requireValueInList(["Please select courier service."])
    .setAllowInvalid(false)
    .build();

  var CourierS = CalSheet.getRange("C7").getValue();

  if (CourierS != "" && Row == 7 && Col == 3 && CourierS === "Shiprocket") {
    CalSheet.getRange("D20").clearContent();
    CalSheet.getRange("C8").clearDataValidations();
    CalSheet.getRange("C8").clearContent();
    CalSheet.getRange("C8").setDataValidation(ShiprocketDD);
    SS.toast("Dropdowns updated successfully", "SUCCESSFUL");

  } else if (CourierS != "" && Row == 7 && Col == 3 && CourierS === "Indiapost") {
    CalSheet.getRange("D20").clearContent();
    CalSheet.getRange("C8").clearDataValidations();
    CalSheet.getRange("C8").clearContent();
    CalSheet.getRange("C8").setDataValidation(IndiapostDD);
    SS.toast("Dropdowns updated successfully", "SUCCESSFUL");

  } else if (CourierS != "" && Row == 7 && Col == 3 && CourierS === "Mail Innovation") {
    CalSheet.getRange("D20").clearContent();
    CalSheet.getRange("C8").clearDataValidations();
    CalSheet.getRange("C8").clearContent();
    CalSheet.getRange("C8").setDataValidation(MailInnoDD);
    SS.toast("Dropdowns updated successfully", "SUCCESSFUL");

  }else if (CourierS == "" && Row == 7 && Col == 3) {
    CalSheet.getRange("D20").clearContent();
    CalSheet.getRange("C8").clearDataValidations();
    CalSheet.getRange("C8").clearContent();
    CalSheet.getRange("C8").setDataValidation(BlankVal);
    SS.toast("Dropdowns updated successfully", "SUCCESSFUL");
  }

}
