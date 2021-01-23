function ImportXml() {
  Logger.clear();
  
//  var NL = String.fromCharCode(10);
//  var QM = String.fromCharCode(34);
  
  var myApp = SpreadsheetApp;
  
  var mySpread = myApp.getActiveSpreadsheet();
  
  // SET SHEETS
  var myImportSheet = mySpread.getSheetByName("query");
//  var myTempSheet = mySpread.getSheetByName("Temp_sheet");
  
  //________________________________________
  // SET RANGE OF ROWS TO RUN PARSER HERE:
  //________________________________________

  for(var i=3; i<=118; i++){   //i: SET ME HERE********  1589//1179//1420//!!929-933 left at 1016
     
//    myTempSheet.clear();
    // var myOCREid = myImportSheet.getRange(i,3).getValue().toString();    
    var myOCRElink = myImportSheet.getRange(i,1).getValue().toString();    
    // var myLink = "http://numismatics.org/ocre/id/" + myOCREid + ".xml";  
    var myLink = myOCRElink + ".xml";  
    var myResponse = UrlFetchApp.fetch(myLink).getContentText();
    var myXML = XmlService.parse(myResponse);
    var myRoot = myXML.getRootElement();
    var myNameSpace = XmlService.getNamespace("http://nomisma.org/nuds");

//    var myElements = myXML.getElement();
    
// obverse: 1,1, 7
//    reverse: 1,1,8
    var myTitle = myRoot.getChild('descMeta',myNameSpace).getChild('title',myNameSpace);
    var myAuthority = myRoot.getChild('descMeta',myNameSpace).getChild('typeDesc',myNameSpace).getChild('authority',myNameSpace).getChild('persname',myNameSpace);
    var myID = myRoot.getChild('control',myNameSpace).getChild('recordId',myNameSpace);
    var myDeity = myRoot.getChild('descMeta',myNameSpace).getChild('typeDesc',myNameSpace).getChild('reverse',myNameSpace).getChild('persname',myNameSpace);
    var myDenomination = myRoot.getChild('descMeta',myNameSpace).getChild('typeDesc',myNameSpace).getChild('denomination',myNameSpace);
    var myManufacture = myRoot.getChild('descMeta',myNameSpace).getChild('typeDesc',myNameSpace).getChild('manufacture',myNameSpace);
    var myMaterial = myRoot.getChild('descMeta',myNameSpace).getChild('typeDesc',myNameSpace).getChild('material',myNameSpace);
    var myMint = myRoot.getChild('descMeta',myNameSpace).getChild('typeDesc',myNameSpace).getChild('geographic',myNameSpace).getChild('geogname',myNameSpace);
    var myObj = myRoot.getChild('descMeta',myNameSpace).getChild('typeDesc',myNameSpace).getChild('objectType',myNameSpace);
    var myPortrait = myRoot.getChild('descMeta',myNameSpace).getChild('typeDesc',myNameSpace).getChild('obverse',myNameSpace).getChild('persname',myNameSpace);
    var myYear = myRoot.getChild('descMeta',myNameSpace).getChild('typeDesc',myNameSpace).getChild('date',myNameSpace);
    if (!(myYear)){
      var myYearFrom = myRoot.getChild('descMeta',myNameSpace).getChild('typeDesc',myNameSpace).getChild('dateRange',myNameSpace).getChild('fromDate',myNameSpace);
      var myYearTo = myRoot.getChild('descMeta',myNameSpace).getChild('typeDesc',myNameSpace).getChild('dateRange',myNameSpace).getChild('toDate',myNameSpace);
      myYearString = myYearFrom.getValue().replace("AD ","") + "|" + myYearTo.getValue().replace("AD ","")
    }
    else{
      myYearString = myYear.getValue().replace("AD ","")
    }
    var myObvLegend = myRoot.getChild('descMeta',myNameSpace).getChild('typeDesc',myNameSpace).getChild('obverse',myNameSpace).getChild('legend',myNameSpace);
    var myRvLegend = myRoot.getChild('descMeta',myNameSpace).getChild('typeDesc',myNameSpace).getChild('reverse',myNameSpace).getChild('legend',myNameSpace);
    var myObvDescription = myRoot.getChild('descMeta',myNameSpace).getChild('typeDesc',myNameSpace).getChild('obverse',myNameSpace).getChild('type',myNameSpace).getChild('description',myNameSpace);
    var myRvDescription = myRoot.getChild('descMeta',myNameSpace).getChild('typeDesc',myNameSpace).getChild('reverse',myNameSpace).getChild('type',myNameSpace).getChild('description',myNameSpace);

//    myTempSheet.clear();
    if (myTitle){
    myImportSheet.getRange(i, 2).setValue(myTitle.getValue());
    }
    if (myID){
    myImportSheet.getRange(i, 3).setValue(myID.getValue());
    }
    if (myAuthority){
    myImportSheet.getRange(i, 4).setValue(myAuthority.getValue());
    }
    if (myDeity){
    myImportSheet.getRange(i, 6).setValue(myDeity.getValue());
    }   
    if (myDenomination){
    myImportSheet.getRange(i, 7).setValue(myDenomination.getValue());
    }
    if (myManufacture){
    myImportSheet.getRange(i, 13).setValue(myManufacture.getValue());
    } 
    if (myMaterial){
    myImportSheet.getRange(i, 14).setValue(myMaterial.getValue());
    } 
    if (myMint){
    myImportSheet.getRange(i, 15).setValue(myMint.getValue());
    }
    if (myObj){
    myImportSheet.getRange(i, 18).setValue(myObj.getValue());
    }  
    if (myPortrait){
    myImportSheet.getRange(i, 19).setValue(myPortrait.getValue());
    } 
     
    myImportSheet.getRange(i, 26).setValue(myYearString);
          
    if (myObvLegend){
    myImportSheet.getRange(i, 16).setValue(myObvLegend.getValue());
    }
    if (myRvLegend){
    myImportSheet.getRange(i, 22).setValue(myRvLegend.getValue());
    }
    if (myObvDescription){
    myImportSheet.getRange(i, 17).setValue(myObvDescription.getValue());
    }
    if (myRvDescription){    
    myImportSheet.getRange(i, 23).setValue(myRvDescription.getValue());
    }
  } // end of for i = m to n 
} //end of function
