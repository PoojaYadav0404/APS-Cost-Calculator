function itemWeightAndPP() {
  // when item name is selected we fetch its weight and purchase price
 
  var SelectedItemSKU = CalSheet.getRange("C4").getValue();
  CalSheet.getRange("C3").setValue(new Date());
  var ItemArr = DDSheet.getRange(2, 1, DDSheet.getLastRow() - 1, 9).getValues();
  var Arr = [];

  if (ActiveSheet.getName() == Cal && Row == 4 && Col == 3 && SelectedItemSKU != "") {

    ItemArr.forEach(function (r) {
      if (r[0] == SelectedItemSKU) {
        Arr.push([r[1],r[8], r[4]]);
      }
    })
    Logger.log(Arr)

    if (Arr.length === 1) {
      CalSheet.getRange("C5").setValue(Arr[0][0]);
      CalSheet.getRange("C9").setValue(Arr[0][1]);
      CalSheet.getRange("C10").setValue(Arr[0][2]);
      SS.toast("Item weight and purchase price updated.", "Message!!")



    } else { CalSheet.getRange("C5").clearContent();
             CalSheet.getRange("C9").clearContent();
             CalSheet.getRange("C10").clearContent();
             var Ui = SpreadsheetApp.getUi();
             Ui.alert("Alert!!", "Either Duplicate entry or values doesn't exist. ", Ui.ButtonSet.OK)

    }
  } else if (ActiveSheet.getName() == Cal && Row == 4 && Col == 3 && SelectedItemSKU == "") {
    CalSheet.getRange("C5").clearContent();
    CalSheet.getRange("C6").clearContent();
    CalSheet.getRange("C7").clearContent();
    CalSheet.getRange("C8").clearContent();
    CalSheet.getRange("C9").clearContent();
    CalSheet.getRange("C10").clearContent();
    CalSheet.getRange("D20").clearContent();
    //SS.toast("Please select Item code/name.", "Alert!!")
  } 
}
