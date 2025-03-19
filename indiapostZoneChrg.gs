function indiapostZoneChrg() {
  //we fetch indiapost prices according to zone from its list as per item weight

  var CurService = CalSheet.getRange("C7").getValue();
  var Zone = CalSheet.getRange("C8").getValue();
  var ItemWeight = CalSheet.getRange("C9").getValue();
  var ArrData = IndiapostSheet.getRange(4, 1, IndiapostSheet.getLastRow() - 3, 3).getValues();

  if (ActiveSheet.getName() == Cal && CurService == "Indiapost" && (Row == 7 || Row == 8) && Col == 3 && CurService != "" && Zone != "" && ItemWeight != "") {
    var Arr = [];

    ArrData.forEach(function (r) {
      if (Zone == r[0] && ItemWeight <= .250) {
        Arr.push(r[1]);

      } else if ( Zone == r[0] && ItemWeight > .250 && ItemWeight <= .5) {
        Arr.push(r[1],r[2])
        
      }else if ( Zone == r[0] && ItemWeight > .5 && ItemWeight <= .75) {
        var MoreX2 = Math.abs(r[2]*2)
        Arr.push(r[1],MoreX2)
        
      }else if ( Zone == r[0] && ItemWeight > .75 && ItemWeight <= 1) {
        var MoreX3 = Math.abs(r[2]*3)
        Arr.push(r[1],MoreX3)
        
      }else if ( Zone == r[0] && ItemWeight > 1 && ItemWeight <= 1.25) {
        var MoreX4 = Math.abs(r[2]*4)
        Arr.push(r[1],MoreX4)
        
      }else if ( Zone == r[0] && ItemWeight > 1.25 && ItemWeight <= 1.5) {
        var MoreX5 = Math.abs(r[2]*5)
        Arr.push(r[1],MoreX5)
        
      }else if ( Zone == r[0] && ItemWeight > 1.5 && ItemWeight <= 1.75) {
        var MoreX6 = Math.abs(r[2]*6)
        Arr.push(r[1],MoreX6)
        
      }else if ( Zone == r[0] && ItemWeight > 1.75 && ItemWeight <= 2) {
        var MoreX7 = Math.abs(r[2]*7)
        Arr.push(r[1],MoreX7)
        
      }else if ( Zone == r[0] && ItemWeight > 2 && ItemWeight <= 2.25) {
        var MoreX8 = Math.abs(r[2]*8)
        Arr.push(r[1],MoreX8)
        
      }else if ( Zone == r[0] && ItemWeight > 2.25 && ItemWeight <= 2.5) {
        var MoreX9 = Math.abs(r[2]*9)
        Arr.push(r[1],MoreX9)
        
      }else if ( Zone == r[0] && ItemWeight > 2.5 && ItemWeight <= 2.75) {
        var MoreX10 = Math.abs(r[2]*10)
        Arr.push(r[1],MoreX10)
        
      }else if ( Zone == r[0] && ItemWeight > 2.75 && ItemWeight <= 3) {
        var MoreX11 = Math.abs(r[2]*11)
        Arr.push(r[1],MoreX11)
        
      }else if ( Zone == r[0] && ItemWeight > 3 && ItemWeight <= 3.25) {
        var MoreX12 = Math.abs(r[2]*12)
        Arr.push(r[1],MoreX12)
        
      }else if ( Zone == r[0] && ItemWeight > 3.25 && ItemWeight <= 3.5) {
        var MoreX13 = Math.abs(r[2]*13)
        Arr.push(r[1],MoreX13)
        
      }else if ( Zone == r[0] && ItemWeight > 3.5 && ItemWeight <= 3.75) {
        var MoreX14 = Math.abs(r[2]*14)
        Arr.push(r[1],MoreX14)
        
      }else if ( Zone == r[0] && ItemWeight > 3.75 && ItemWeight <= 4) {
        var MoreX15 = Math.abs(r[2]*15)
        Arr.push(r[1],MoreX15)
        
      }else if ( Zone == r[0] && ItemWeight > 4 && ItemWeight <= 4.25) {
        var MoreX16 = Math.abs(r[2]*16)
        Arr.push(r[1],MoreX16)
        
      }else if ( Zone == r[0] && ItemWeight > 4.25 && ItemWeight <= 4.5) {
        var MoreX17 = Math.abs(r[2]*17)
        Arr.push(r[1],MoreX17)
        
      }else if ( Zone == r[0] && ItemWeight > 4.5 && ItemWeight <= 4.75) {
        var MoreX18 = Math.abs(r[2]*18)
        Arr.push(r[1],MoreX18)
        
      }else if ( Zone == r[0] && ItemWeight > 4.75 && ItemWeight <= 5) {
        var MoreX19 = Math.abs(r[2]*19)
        Arr.push(r[1],MoreX19)
        
      }else if ( Zone == r[0] && ItemWeight > 5 && ItemWeight <= 5.25) {
        var MoreX20 = Math.abs(r[2]*20)
        Arr.push(r[1],MoreX20)
        
      }else if ( Zone == r[0] && ItemWeight > 5.25 && ItemWeight <= 5.5) {
        var MoreX21 = Math.abs(r[2]*21)
        Arr.push(r[1],MoreX21)
        
      }else if ( Zone == r[0] && ItemWeight > 5.5 && ItemWeight <= 5.75) {
        var MoreX22 = Math.abs(r[2]*22)
        Arr.push(r[1],MoreX22)
        
      }else if ( Zone == r[0] && ItemWeight > 5.75 && ItemWeight <= 6) {
        var MoreX23 = Math.abs(r[2]*23)
        Arr.push(r[1],MoreX23)
        
      }else if ( Zone == r[0] && ItemWeight > 6 && ItemWeight <= 6.25) {
        var MoreX24 = Math.abs(r[2]*24)
        Arr.push(r[1],MoreX24)
        
      }else if ( Zone == r[0] && ItemWeight > 6.25 && ItemWeight <= 6.5) {
        var MoreX25 = Math.abs(r[2]*25)
        Arr.push(r[1],MoreX25)
        
      }else if ( Zone == r[0] && ItemWeight > 6.5 && ItemWeight <= 6.75) {
        var MoreX26 = Math.abs(r[2]*26)
        Arr.push(r[1],MoreX26)
        
      }else if ( Zone == r[0] && ItemWeight > 6.75 && ItemWeight <= 7) {
        var MoreX27 = Math.abs(r[2]*27)
        Arr.push(r[1],MoreX27)
        
      }else if ( Zone == r[0] && ItemWeight > 7 && ItemWeight <= 7.25) {
        var MoreX28 = Math.abs(r[2]*28)
        Arr.push(r[1],MoreX28)
        
      }else if ( Zone == r[0] && ItemWeight > 7.25 && ItemWeight <= 7.5) {
        var MoreX29 = Math.abs(r[2]*29)
        Arr.push(r[1],MoreX29)
        
      }else if ( Zone == r[0] && ItemWeight > 7.5 && ItemWeight <= 7.75) {
        var MoreX30 = Math.abs(r[2]*30)
        Arr.push(r[1],MoreX30)
        
      }else if ( Zone == r[0] && ItemWeight > 7.75 && ItemWeight <= 8) {
        var MoreX31 = Math.abs(r[2]*31)
        Arr.push(r[1],MoreX31)
        
      }else if ( Zone == r[0] && ItemWeight > 8 && ItemWeight <= 8.25) {
        var MoreX32 = Math.abs(r[2]*32)
        Arr.push(r[1],MoreX32);
        
      }else if ( Zone == r[0] && ItemWeight > 8.25 && ItemWeight <= 8.5) {
        var MoreX33 = Math.abs(r[2]*33)
        Arr.push(r[1],MoreX33);
        //................aftrer that calculations nedded
        
      }else if ( Zone == r[0] && ItemWeight > 8.5 && ItemWeight <= 8.75) {
        var MoreX34 = Math.abs(r[2]*34)
        Arr.push(r[1],MoreX34);
        
      }else if ( Zone == r[0] && ItemWeight > 8.75 && ItemWeight <= 9) {
        var MoreX35 = Math.abs(r[2]*35)
        Arr.push(r[1],MoreX35);
        
      }else if ( Zone == r[0] && ItemWeight > 9 && ItemWeight <= 9.25) {
        var MoreX36 = Math.abs(r[2]*36)
        Arr.push(r[1],MoreX36);
        
      }else if ( Zone == r[0] && ItemWeight > 9.25 && ItemWeight <= 9.5) {
        var MoreX37 = Math.abs(r[2]*37)
        Arr.push(r[1],MoreX37);
        
      }else if ( Zone == r[0] && ItemWeight > 9.5 && ItemWeight <= 9.75) {
        var MoreX38 = Math.abs(r[2]*38)
        Arr.push(r[1],MoreX38);
        
      }else if ( Zone == r[0] && ItemWeight > 9.75 && ItemWeight <= 10) {
        var MoreX39 = Math.abs(r[2]*39)
        Arr.push(r[1],MoreX39);
        
      }else if ( Zone == r[0] && ItemWeight > 10 && ItemWeight <= 10.25) {
        var MoreX40 = Math.abs(r[2]*40)
        Arr.push(r[1],MoreX40);
        
      }else if ( Zone == r[0] && ItemWeight > 10.25 && ItemWeight <= 10.5) {
        var MoreX41 = Math.abs(r[2]*41)
        Arr.push(r[1],MoreX41);
        
      }else if ( Zone == r[0] && ItemWeight > 10.5 && ItemWeight <= 10.75) {
        var MoreX42 = Math.abs(r[2]*42)
        Arr.push(r[1],MoreX42);
        
      }else if ( Zone == r[0] && ItemWeight > 10.75 && ItemWeight <= 11) {
        var MoreX43 = Math.abs(r[2]*43)
        Arr.push(r[1],MoreX43);
        
      }else if ( Zone == r[0] &&  ItemWeight < 11) {
        SS.toast( "Contact automation team to set formula for more than 11 kg weight only for India Post. ", "Alert!!")}
      
    })
    Logger.log(Arr.length)

    if (Arr.length === 1 && ItemWeight <= .250) {
      CalSheet.getRange("D20").setValue(Arr);

    } else if (Arr.length === 2 && ItemWeight > .250 && ItemWeight <= 11) {
      let Sum = 0;

      Arr.forEach(item => {
          Sum += item;
      });

      //console.log(Sum);
      
      CalSheet.getRange("D20").setValue(Sum);

    } else if (Arr.length === 0 || Arr.length > 2){ 
      CalSheet.getRange("D20").clearContent();
      var Ui = SpreadsheetApp.getUi();
      Ui.alert("Alert!!", "Either Duplicate entry or values doesn't exist. ", Ui.ButtonSet.OK)

    }

  } 

}
