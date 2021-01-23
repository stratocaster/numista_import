function ListImagesTempSheet() {

  var myApp = SpreadsheetApp;
  
  var mySpread = myApp.getActiveSpreadsheet();
  var myLinksSheet = mySpread.getSheetByName("Twin_Images");
  var myTempSheet = mySpread.getSheetByName("Temp_sheet");
  myTempSheet.clear();
  
//  var myAvFolder = DriveApp.getFolderById('1-vXKBtgGg6JJRXv4i8IMhVrOL3m6OZT6'); // replace FOLDER-ID with your folder's ID
//  var myRvFolder = DriveApp.getFolderById('1zMhf-Nw-ZUO5auaaLS7FBOXrS4lCjsDF');
  var myAvFolder = DriveApp.getFolderById('1e3ltW6WnSVFR9SuEbT2ojokenXF6xBRx'); // replace FOLDER-ID with your folder's ID
  var myRvFolder = DriveApp.getFolderById('1LPzICQHZ-GKRSOMCgP99K8ZlYlRK9waU');
    
  var results = [];
  // list all pdf files in the folder
  var myAvJpgs = myAvFolder.getFilesByType(MimeType.JPEG);
  var myRvJpgs = myRvFolder.getFilesByType(MimeType.JPEG);  
  
  var myLinkArray = new Array;
  var myNameArray = new Array;
  var myTDarray = new Array;
  // AV loop through found files in the folder
    while (myAvJpgs.hasNext()) { /// TURN ON     
      var myfile = myAvJpgs.next();     
      var fname = myfile.getName().toLowerCase();    
      var fID = myfile.getId();    
      myLinkArray.push("https://drive.google.com/uc?id="+fID);
      myNameArray.push(fname);
  }
  
  for (var i=0; i<myLinkArray.length; i++){
    var myOneElement = new Array (2);
    myOneElement[1] = myLinkArray[i];
    myOneElement[0] = myNameArray[i];
    myTDarray[i] = myOneElement;
  }
  
  var myCells = myTempSheet.getRange(3,1,myLinkArray.length,2);
  myCells.setValues(myTDarray);
  
  myLinkArray = [];
  myTDarray = [];
  myNameArray = [];
  // RV loop through found files in the folder
  while (myRvJpgs.hasNext()) { /// TURN ON     
      var myfile = myRvJpgs.next();     
      var fname = myfile.getName().toLowerCase();    
      var fID = myfile.getId();    
      myLinkArray.push("https://drive.google.com/uc?id="+fID);
      myNameArray.push(fname);
  }
    
  for (var i=0; i<myLinkArray.length; i++){
    var myOneElement = new Array (1);
    myOneElement[1] = myLinkArray[i];
    myOneElement[0] = myNameArray[i];
    myTDarray[i] = myOneElement;
  }
  
  var myCells = myTempSheet.getRange(3,3,myLinkArray.length,2);
  myCells.setValues(myTDarray);
}
