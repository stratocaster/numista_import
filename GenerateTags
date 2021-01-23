function myTags () {
  
//  var NL = String.fromCharCode(10);
//  var QM = String.fromCharCode(34);
  
  var myApp = SpreadsheetApp;
  var mySpread = myApp.getActiveSpreadsheet();
  
  // SET SHEETS
  //var mySheet = mySpread.getActiveSheet();
  var myImportSheet = mySpread.getSheetByName("query");
  var myMapSheet = mySpread.getSheetByName("MAP_Tags");
      
  var myID = 0;
  var myTriggers = "";
  var myTriggerArray = [];
  var myDescription = "";
  var myTagsArray = "";
    
  for (var myPos = 3; myPos<=myImportSheet.getLastRow(); myPos++){
    myDescription = myImportSheet.getRange(myPos,17).getValue.toString();
    myDescription = myDescription + myImportSheet.getRange(myPos,22).getValue.toString();
    myDescription = myDescription.toLowerCase();
    for (var i=2; i<=myMapSheet.getLastRow(); i++){
      myTriggers = myMapSheet.getRange(i,4).getValue.toString();
      if (myTriggers !== ""){
        myTriggerArray = myTriggers.split(", ");
        myTagsArray = "";
        for (var j=0; j<myTriggerArray.length; j++){
          if (myDescription.indexOf(myTriggerArray[j])>=0){
              myID = myMapSheet.getRange(i,2).getValue;
            if (myTagsArray !== ""){
              myTagsArray = myTagsArray + ", ";
              break;
            }
            myTagsArray = myTagsArray + myID;
              }
        }
        if (myTagsArray !== ""){
          var myCell = myImportSheet.getRange(myPos,40);
          myTagsArray = "23, 15";
          myCell.setValue(myTagsArray)
        }
        else{
          myTagsArray = "0, 0";
          myCell.setValue(myTagsArray)
        }
      }      
    }            
  }      
}
