//@NotOnlyCurrentDoc

function myImgLinks6 () {

  Logger.clear();
  
  var NL = String.fromCharCode(10);
  var QM = String.fromCharCode(34);
  
  var myApp = SpreadsheetApp;
  
  var mySpread = myApp.getActiveSpreadsheet();
  
  // SET SHEETS
  //var mySheet = mySpread.getActiveSheet();
  var myImportSheet = mySpread.getSheetByName("query");
  var myTempSheet = mySpread.getSheetByName("Temp_sheet");
//  var myLinksSheet = mySpread.getSheetByName("Atom_Links");
  var myTwinImgSheet = mySpread.getSheetByName("Twin_Images");
  
  //________________________________________
  // SET RANGE OF ROWS TO RUN PARSER HERE:
  //________________________________________
 
//  var myT = 28;
  for(var i=40; i<=41; i++){   //i: SET ME HERE********
//  Logger.log("Step ",i," starts here");
      
    //////////////////////////////////////////////////////
    //Import all Museum samples in Temp sheet
    //////////////////////////////////////////////////////    
    
    var myOCREid = myImportSheet.getRange(i,3).getValue().toString();     
    var myLink = "http://nomisma.org/query?query=PREFIX+rdf%3A%09%09%3Chttp%3A%2F%2Fwww.w3.org%2F1999%2F02%2F22-rdf-syntax-ns%23%3E%0D%0APREFIX+dcterms%3A%09%09%3Chttp%3A%2F%2Fpurl.org%2Fdc%2Fterms%2F%3E%0D%0APREFIX+nm%3A%09%09%3Chttp%3A%2F%2Fnomisma.org%2Fid%2F%3E%0D%0APREFIX+nmo%3A%09%09%3Chttp%3A%2F%2Fnomisma.org%2Fontology%23%3E%0D%0APREFIX+foaf%3A%09%09%3Chttp%3A%2F%2Fxmlns.com%2Ffoaf%2F0.1%2F%3E%0D%0APREFIX+skos%3A%09%3Chttp%3A%2F%2Fwww.w3.org%2F2004%2F02%2Fskos%2Fcore%23%3E%0D%0A%0D%0ASELECT+%3Fobject+%3Fidentifier+%3Fdiameter+%3Fweight+%3Faxis+%3Fcollection+%3FobvRef+%3FrevRef+%3FcomRef+WHERE+%7B%0D%0A%09%7B%3Fobject+nmo%3AhasTypeSeriesItem+%3Chttp%3A%2F%2Fnumismatics.org%2Focre%2Fid%2F" + myOCREid + "%3E+%7D%0D%0A%09%3Fobject+rdf%3Atype+nmo%3ANumismaticObject+.%0D%0A%09OPTIONAL+%7B+%3Fobject+nmo%3AhasWeight+%3Fweight+%7D%0D%0A%09OPTIONAL+%7B+%3Fobject+nmo%3AhasDiameter+%3Fdiameter+%7D%0D%0A%09OPTIONAL+%7B+%3Fobject+nmo%3AhasAxis+%3Faxis+%7D%0D%0A%09OPTIONAL+%7B+%3Fobject+dcterms%3Aidentifier+%3Fidentifier+%7D%0D%0A%09OPTIONAL+%7B+%3Fobject+nmo%3AhasCollection+%3FcolUri+.%0D%0A%09%09%3FcolUri+skos%3AprefLabel+%3Fcollection+FILTER%28langMatches%28lang%28%3Fcollection%29%2C+%22EN%22%29%29%7D%0D%0A%09OPTIONAL+%7B+%3Fobject+foaf%3Adepiction+%3FcomRef+%7D%0D%0A%09OPTIONAL+%7B+%3Fobject+nmo%3AhasObverse+%3Fobverse+.%0D%0A%09%09%3Fobverse+foaf%3Adepiction+%3FobvRef+%7D%0D%0A%09OPTIONAL+%7B+%3Fobject+nmo%3AhasReverse+%3Freverse+.%0D%0A%09%09%3Freverse+foaf%3Adepiction+%3FrevRef+%7D%0D%0A%7D&output=csv";    
    var myResponse = UrlFetchApp.fetch(myLink);
    var myResponseData = myResponse.getBlob().getDataAsString();
    var myData = Utilities.parseCsv(myResponseData, ',');
    myTempSheet.clear();
    myTempSheet.getRange(1, 1, myData.length, myData[0].length).setValues(myData);  
      
    /////////////////////
    //COMPUTE WEIGHT and DIAMETER averages
    ////////////////////////
    var myTmass = 0;
    var myCmass = 0;
    var myTdiam = 0;
    var myCdiam = 0;
    var myMinDiam = 1000;
    var myMinMass = 1000000;
    var myMaxMass = -1;
    var myMaxDiam = -1;
    var myVariance = 0;
    var myAxes = [];
    var mySumAxes = 0;
    var mySumSin = 0;
    var mySumCos = 0;
    
    var myOVlink = "";
      var myRVlink = "";
      var myCreditType = "";
      var myImgCredits ="";
      var myTwinLink = "";
      var myWebLink ="";
    
      var myPermission = 0; // permission: 2=good; 1=not the best; 0=bad;
      var myTmpPermission = 0;
    
      ///GO ROW BY ROW through the TEMP SHEET --- AXIS; DIAMETER, MASS
    for (var j=2; j<=myTempSheet.getLastRow(); j++){
      //mass
      var myTC = myTempSheet.getRange(j,4).getValue();      
      if (myTC !== "") {
      myTmass += myTC;
        if (myTC > myMaxMass){
          myMaxMass = myTC;
        }
        if (myTC < myMinMass){
          myMinMass = myTC;
        }
      myCmass += 1;
      }
      
      //diameter
      var myTC = myTempSheet.getRange(j,3).getValue();      
      if (myTC !== "") {
        myTdiam += myTC;
        if (myTC > myMaxDiam){
          myMaxDiam =myTC;
        }
        if (myTC < myMinDiam){
          myMinDiam = myTC;
        }
        myCdiam += 1;
      }
      
      //axis //coin // medal // variable // unknown
      var myTC = myTempSheet.getRange(j,5).getValue();      
      if (myTC !== "") {
        myAxes.push(myTC);
        mySumAxes += myTC;
        mySumSin += Math.sin(myTC*Math.PI/6);
        mySumCos += Math.cos(myTC*Math.PI/6);
      }        
    
//COMPUTE Image links //personnel //site_autorise //autre
      var myChange = false;
      if (myPermission <= 4){
        var myTmpImgCredits = myTempSheet.getRange(j,6).getValue();
//        var myCreditType ="";
      switch (myTmpImgCredits)  { 
        case "American Numismatic Society": //*****************=========================     =OK
          myTmpPermission = 5;
          if (myTmpPermission > myPermission){
            myChange = true;
            myImgCredits = "American Numismatic Society (ANS)";
            myCreditType = "site_autorise";
            myPermission = myTmpPermission;
            }
          break    
        case "Thuringian Museum for Pre- and Early History": //*****************=========================     =OK
          myTmpPermission = 5;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Museum für Ur- und Frühgeschichte Thüringens";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break
        case "State Museum of Prehistory Halle": //*****************=========================     =OK
          myTmpPermission = 5;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Landesmuseum für Vorgeschichte Halle";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break         
        case "Kunstmuseum Moritzburg Halle (Saale)": //??????????????????????????????
          myTmpPermission = 5;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Kulturstiftung Sachsen-Anhalt - Kunstmuseum Moritzburg Halle (Saale)";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break          
        case "Oldenburg Municipal Museum": //*****************=========================     =OK
          myTmpPermission = 5;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Stadtmuseum Oldenburg";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break
        case "State Coin Collection of Munich": //*****************=========================     =OK
          myTmpPermission = 5;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Staatliche Münzsammlung München";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break
        case "Münzkabinett der Universität Göttingen": //*****************=========================     =OK
          myTmpPermission = 5;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Münzkabinett der Universität Göttingen";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break        
        // case "Münzkabinett der Staatlichen Museen zu Berlin": //*****************
        //   myTmpPermission = 1;
        //   if (myTmpPermission > myPermission){
        //      myChange = true;
        //      myImgCredits = "Münzkabinett, Staatliche Museen zu Berlin (CC BY-NC-SA)";
        //      myCreditType = "site_autorise";
        //      myPermission = myTmpPermission;
        //     }
        //   break;    
        // case "Münzkabinett Berlin": //*****************=========================     =OK
        //   myTmpPermission = 1;
        //   if (myTmpPermission > myPermission){
        //      myChange = true;
        //      myImgCredits = "Münzkabinett, Staatliche Museen zu Berlin (CC BY-NC-SA)";
        //      myCreditType = "site_autorise";
        //      myPermission = myTmpPermission;
        //     }
        //   break;
        case "British Museum": //*****************
          myTmpPermission = 2;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Trustees of the British Museum";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break;
        case "The Trustees of the British Museum": //*****************
          myTmpPermission = 2;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Trustees of the British Museum";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break;          
        case "Bibliothèque nationale de France": //***************** 30,908 ????
          myTmpPermission = 5;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Bibliothèque nationale de France";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break;
        case "Furman University Libraries":   // ????
          myTmpPermission = 5;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Furman University Libraries Coin Collection (ODC-ODbL)";
             myCreditType = "autre";
             myPermission = myTmpPermission;
            }
          break;
        case "University of Graz":  //*****************2,160  //*****************=========================     =OK GREEN IMAGES
          myTmpPermission = 3;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Institute of Classics/University of Graz";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break;
        case "Institute for Advanced Technology in the Humanites":  //????
          myTmpPermission = 5;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Institute for Advanced Technology in the Humanites (ODC-ODbL)";
             myCreditType = "autre";
             myPermission = myTmpPermission;
            }
          break;
        case "Institute of Archaeology, University of Warsaw": //*****************7,210   //???
          myTmpPermission = 5;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Institute of Archaeology, University of Warsaw";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break;
        case "Museu Arqueològic de Llíria": //*****************5,990 *=========================     =OK
          myTmpPermission = 4;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Museu de Prehistòria de València";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break;
        case "Museu de Prehistòria de València": //*****************5,990 *=========================     =OK
           myTmpPermission = 4;
           if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Museu de Prehistòria de València";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break;
        // case "Münzsammlung des Seminars für Alte Geschichte der Albert-Ludwigs-Universität": //***************** 10,542 =========================     =OK
        //   myTmpPermission = 5;
        //   if (myTmpPermission > myPermission){
        //      myChange = true;
        //      myImgCredits = "Münzsammlung des Seminars für Alte Geschichte, Albert-Ludwigs-Universität Freiburg";
        //      myCreditType = "site_autorise";
        //      myPermission = myTmpPermission;
        //     }
        //   break;
        case "Open Context": //***************** 174   ???
          myTmpPermission = 5;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Open Context (CC BY)";
             myCreditType = "autre";
             myPermission = myTmpPermission;
            }
          break;
        case "Fitzwilliam Museum": //***************** 5337
          myTmpPermission = 4;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "The Fitzwilliam Museum, Cambridge";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break;          
        case "University of Oxford": //***************** 5337    ???
          myTmpPermission = 5;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Heberden Coin Room, Ashmolean Museum, University of Oxford";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break; 
//        case "Römisch Germanische Kommission (RGK) - Germany": //*****************1,457
//          myImgCredits = "Römisch Germanische Kommission";
//          myCreditType = "site_autorise";
//          break;
        case "The Metropolitan Museum of Art": //***************** 54  //???
          myTmpPermission = 5;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "The Metropolitan Museum of Art (CC 0)";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break;
        case "University College Dublin": //***************** 262   ///???
          myTmpPermission = 5;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "University College Dublin (CC BY-SA)";
             myCreditType = "site_autorise";
             myPermission = myTmpPermission;
            }
          break;  
        case "Universität Tübingen": //***************** 541###########  ///???
          myTmpPermission = 5;
          if (myTmpPermission > myPermission){
             myChange = true;
             myImgCredits = "Institut für Klassische Archäologie der Universität Tübingen";
             myCreditType = "autre";
             myPermission = myTmpPermission;
            }
          break;
//        case "Classical Numismatic Group": //*****************
//          myImgCredits = "Classical Numismatic Group, Inc.";
//          myCreditType = "site_autorise";
//          break;
//        case "Heritage Auctions": //*****************
//          myImgCredits = "Heritage Auctions";
//          myCreditType = "site_autorise";
//          break;
//        case "CGB": //*****************
//          myImgCredits = "CGB";
//          myCreditType = "site_autorise";
//          break;
//        case "Roma Numismatics": //*****************
//          myImgCredits = "Roma Numismatics Limited";
//          myCreditType = "site_autorise";
//          break;
//        case "Numismatica Ars Classica": //*****************
//          myImgCredits = "Numismatica Ars Classica NAC AG";
//          myCreditType = "site_autorise";
//          break;
//        case "Hess Divo": //*****************
//          myImgCredits = "Hess Divo";
//          myCreditType = "site_autorise";
//          break;
//Numismatik Naumann GmbH          
//        default:
//          myPermission = 0;
//          myImgCredits = "";
//          break;
      } // end of switch
//          
          if (myChange == true){
             myWebLink = myTempSheet.getRange(j,1).getValue();
             myOVlink = myTempSheet.getRange(j,7).getValue(); 
             if (myOVlink == ""){
              myTwinLink = myTempSheet.getRange(j,9).getValue();
              if (myTwinLink==""){
                 myPermission = 0;
                 myImgCredits = "";
                 myCreditType = "";
                 myWebLink = "";
               }
             }
             else {
               myRVlink = myTempSheet.getRange(j,8).getValue();
               if (myRVlink==""){
                 myPermission = 0;
                 myOVlink = "";
                 myImgCredits = "";
                 myCreditType = "";
                 myWebLink = ""
               }
//               if (myImgCredits=="Trustees of the British Museum"){
//                 myPermission = false
//                 myOVlink = "";
//                 myRVlink = "";
//                 myImgCredits = "";
//                 myCreditType = "";
//                 myWebLink = ""
//               }
             }
           }
        }  // end of if myPermission is false        
    } // end of for j in temp sheet
    //Logger.log("hello");
    
    if ((myPermission >= 1)){
      if (myPermission == 2){ //if image needs dual link
        if (myTwinLink !== ""){ //if dual link exists
          myCell = myTwinImgSheet.getRange(i,1);           
          myCell.setValue(myTwinLink);                         
          //extract image name in cell 2           
          myCell = myTwinImgSheet.getRange(i,2);      
          var myLinkArray = [{}];      
          myLinkArray = myTwinLink.toString().split("/");       
          myCell.setValue(myLinkArray[myLinkArray.length-1])
          myCell = myTwinImgSheet.getRange(i,6);
          myCell.setValue(myImgCredits);
          myCell = myTwinImgSheet.getRange(i,7);
          myCell.setValue(myCreditType);               
          myCell = myTwinImgSheet.getRange(i,3);
          myCell.setValue(myWebLink);
        }
      }
      else{
        myCell = myTwinImgSheet.getRange(i,4);
        myCell.setValue(myOVlink);        
        myCell = myTwinImgSheet.getRange(i,5);
        myCell.setValue(myRVlink);   
        myCell = myTwinImgSheet.getRange(i,6);
        myCell.setValue(myImgCredits);
        myCell = myTwinImgSheet.getRange(i,7);
        myCell.setValue(myCreditType);               
        myCell = myTwinImgSheet.getRange(i,3);
        myCell.setValue(myWebLink);
      } //end my Twin Links if
    } // end of if MyPermission
    
    //set mass rounded to the nearest .1
          if (myTmass !== 0){
            myTmass = Math.round(10*myTmass / myCmass)/10; 
            myCell = myImportSheet.getRange(i,36);
            myCell.setValue(myTmass);
          if ((myCmass >= 2)&&(myMinMass !== myMaxMass)){ 
            myCell = myImportSheet.getRange(i,38);
            myCell.setValue(myMinMass + "–" + myMaxMass + " g;");
    }
    }
    //set average diameter rounded to the nearest .5
    if (myTdiam !== 0){
    myTdiam = Math.round(2*myTdiam / myCdiam)/2;
      myCell = myImportSheet.getRange(i,37);
      myCell.setValue(myTdiam);
      if ((myCdiam >= 2)&&(myMinDiam !== myMaxDiam)){ 
      myCell = myImportSheet.getRange(i,39);
      myCell.setValue(myMinDiam + "–" + myMaxDiam + " mm;");
    }
    } 
    //set AXES
    var myCaxes = myAxes.length;
    var myVariance = 0;
    var myAxisN = "";
    var myAverageAxis = Math.atan2(mySumSin/myCaxes, mySumCos/myCaxes);
    myAverageAxis = myAverageAxis*6/Math.PI;
    if (myAverageAxis<0){
      myAverageAxis +=12
    }
//    mySumAxes = mySumAxes/myCaxes;
    
    if (myCaxes>1){
      for (var myA=0; myA<myCaxes; myA++){
        myVariance = myVariance + Math.min(Math.abs(myAverageAxis-myAxes[myA]), Math.abs(Math.abs(myAverageAxis-myAxes[myA])-12))
      }
//      myCell = myImportSheet.getRange(i,40);
//      myCell.setValue(myVariance);      
    myVariance = myVariance / myCaxes;
    if (myCaxes > 4){
      if (myVariance<0.35){
        if (Math.abs(myAverageAxis-6) <1){
          myAxisN = "coin";
        }
        if (Math.min(Math.abs(myAverageAxis),Math.abs(myAverageAxis-12)) <1){
          myAxisN = "medal";
        }
      }
    
    if (myVariance>=0.5){
          myAxisN = "variable";
        }
      myCell = myImportSheet.getRange(i,40);
      myCell.setValue(myAxisN);      
    } 
      }
    
  } // end of for i = m to n 
} //end of function
