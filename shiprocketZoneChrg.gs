function shiprocketZoneChrg() {
  //we fetch shiprocket prices according to zone from its list as per item weight

  var CurService = CalSheet.getRange("C7").getValue();
  var Zone = CalSheet.getRange("C8").getValue();
  var ItemWeight = CalSheet.getRange("C9").getValue();
  var ZoneDataRange = ShiprocketSheet.getRange(1, 1, 1, ShiprocketSheet.getLastColumn()).getValues()[0];
  var WeightDataRange = ShiprocketSheet.getRange(1, 1, ShiprocketSheet.getRange("A1").getDataRegion().getLastRow(), ShiprocketSheet.getLastColumn()).getValues();

  if (ActiveSheet.getName() == Cal && CurService == "Shiprocket" && (Row == 7 || Row == 8) && Col == 3 && CurService != "" && Zone != "" && ItemWeight != "") {

    //finding index of zone(like USA, Uk)
    var ZoneIndex = ZoneDataRange.indexOf(Zone);

    var Arr = [];
    WeightDataRange.forEach(function (r) {
      if (r[0] == ItemWeight) {
        Arr.push(r[ZoneIndex])

      }
    })
    if (Arr.length === 1) {
      CalSheet.getRange("D20").setValue(Arr);
      //Logger.log(Arr);
    } else {
      CalSheet.getRange("D20").clearContent();
      var Ui = SpreadsheetApp.getUi();
      Ui.alert("Alert!!", "Either Duplicate entry or values doesn't exist. ", Ui.ButtonSet.OK)
    }


  } else if (ActiveSheet.getName() == Cal && (Row == 7 || Row == 8) && Col == 3 && CurService == "" && Zone == "") {
    CalSheet.getRange("D20").clearContent();
  }

}
