/*
  created by: Nicolas Botello
  Any concerns or requests https://github.com/TAMUSHPE/mettingMacroSheets/issues
 */
//Global Variable 
//Magical word not to use first event column
var noFirstEventColumn = "NONE";
//sets up custom menu to bring up our options
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Custom Menu')
      .addItem('Points Meeting Macro', 'showSidebar')
      .addToUi();
}

/**
 * sheet returns a an object that contains values needed to compare data
 * @param  {[object]}  sheet        [sheet object to control excel]
 * @param  {[string]}  FNC          [first name column letter]
 * @param  {[Integer]} FNR          [first name int]
 * @param  {[string]}  LNC          [Last name column]
 * @return {[object]}  customSheet  [sheet object containing values needed]
 */
function sheet(sheetObj, FNC,FNR,LNC)
{
  var sheet =
  {
    sheet: sheetObj,
    firstNameColumn: letterToColumn(FNC), //conver to number from letter
    LastNameColumn: letterToColumn(LNC),
    firstNameRow: FNR,
    totalRows: (sheetObj.getLastRow()- FNR)+1, //total rows that we will be using
    lastRow: sheetObj.getLastRow()
  };
  //get all values used
  //(startRow, startColumn, numRows, numColumns
  //get all first name values
  sheet.firstNameValues= sheet.sheet.getSheetValues(sheet.firstNameRow, sheet.firstNameColumn
    , sheet.totalRows ,1);
  sheet.lastNameValues = sheet.sheet.getSheetValues(sheet.firstNameRow, sheet.LastNameColumn
    , sheet.totalRows ,1);
   return sheet;
}
/**
 * setupRosterSheet sets up other columns needed by roster sheet
 * @param  {[object]} sheet                [sheet object]
 * @param  {[String]} tamuApplicantColumn  [string that represents the column that contains tamu applicant]
 * @param  {[String]} nationalMemberColumn [string that represents the column that contains national member]
 */
function setupRosterSheet(sheet,tamuApplicantColumn,nationalMemberColumn){

  sheet.tamuApplicantColumn = letterToColumn(tamuApplicantColumn);
  sheet.nationalMemberColumn = letterToColumn(nationalMemberColumn);
  //(startRow, startColumn, numRows, numColumns
  //get all tamuApplicant data
  sheet.tamuApplicantValues= sheet.sheet.getSheetValues(sheet.firstNameRow, sheet.tamuApplicantColumn
    , sheet.totalRows ,1);
  sheet.nationalMemberValues= sheet.sheet.getSheetValues(sheet.firstNameRow, sheet.nationalMemberColumn
    , sheet.totalRows ,1);
}
/**
 * setupMeetingSheet setus up shirt colum needed to give extra points
 * @param  {[type]} sheet         [sheet object]
 * @param  {[String]} shirtColumn [string that represents the column that contains shirt values]
 * @param  {[String} firstEvent   [string that represents the column that contains first shpe event?]
 */
function setupMeetingSheet(sheet,shirtColumn, firstEvent)
{
  sheet.shirtColumn = letterToColumn(shirtColumn);
  //if magical word is there then don't check the event column
  if(firstEvent !== noFirstEventColumn)
  {
    sheet.firstEventColumn = letterToColumn(firstEvent);
      sheet.firstEventValues =  sheet.sheet.getSheetValues(sheet.firstNameRow, sheet.firstEventColumn
    , sheet.totalRows ,1);
  }
  else
  {
    sheet.firstEventColumn = noFirstEventColumn;
  }
  sheet.shirtColumnValues =  sheet.sheet.getSheetValues(sheet.firstNameRow, sheet.shirtColumn
    , sheet.totalRows ,1);
}
/**
 * compareSheet compares the two sheets current sheet is the "Roster" and pastSheet is the url sheet and sets points
 * pointsFunction is the function used to evaluate the points it takes in two params returns a Number the points
 *                @param {Object} [pastSheet] [object that contains helper values for pastSheet]
 *                @param {Number} [i]         [the index of the current row it is on]
 * extraSetup is a param where you can pass a function and it will call it add more fields to use in pointsFunction params:
 *                @param {Object} [pastSheet] [object that contains helper values for pastSheet]
 * @param  {object} data           [object that contains all the fields that are needed to setup the sheets can be found in retrieveFields function]
 * @param  {function} pointsFunction [Function used to give assign the points ]
 * @param  {function} extraSetup     [Function can be empty, possible to help setup other fields needed to calculate points]
 */
function compareSheet(data, pointsFunction, extraSetup)
{
  var url = data.url;
  var targetColumn = data.pointsColumn;
  //points Function must be a function or everything will break
  if(typeof pointsFunction !== "function")
  {
    return;
  } 
  var FirstSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var SecondSheet = SpreadsheetApp.openByUrl(url);
  var currentSheet = sheet(FirstSheet, data.currentSheet.firstName ,3, data.currentSheet.lastName);
  currentSheet.pointValues = currentSheet.sheet.getSheetValues(currentSheet.firstNameRow, letterToColumn(targetColumn), 
   currentSheet.totalRows,1);
  setupRosterSheet(currentSheet, data.currentSheet.tamuApplicant, data.currentSheet.nationalMember);

  var pastSheet = sheet(SecondSheet, data.pastSheet.firstName ,2, data.pastSheet.lastName);
  setupMeetingSheet(pastSheet, data.pastSheet.shirt, data.pastSheet.firstEvent);
  //add a possible setup function
  if(typeof extraSetup !== "undefined" && typeof extraSetup === "function")
  {
    extraSetup(pastSheet);
  }
  for(var i = 0; i < pastSheet.firstNameValues.length; i++)
  {
    var found = false;
    for(var j = 0; j < currentSheet.firstNameValues.length; j++)
    {
      
      if(isMembershipComplete(currentSheet,j))
      {
         //cell from all names part of org
        if(compareNames(pastSheet,i, currentSheet,j))
        {                       
          currentSheet.pointValues[j][0]= pointsFunction(pastSheet, i);
          found = true;
          break;
        }
      }
    }
    //if this is your first SHPE meeting yes then highlight red
    //only if its not equal to the magical word for no first event column check
    if ( pastSheet.firstEventColumn !== noFirstEventColumn && String(pastSheet.firstEventValues[i][0]).toLowerCase().trim() === "yes")
    {
      pastSheet.sheet.getRange("A"+(i+2)+":"+"I"+(i+2)).setBackground("red");
    }
    //highlight names that have not been found so human can double check
    else if(!found)
    {
      pastSheet.sheet.getRange("A"+(i+2)+":"+"I"+(i+2)).setBackground("yellow");
    }
  }  
  //change all values
  currentSheet.sheet.getRange(targetColumn+currentSheet.firstNameRow+":"+targetColumn+currentSheet.lastRow).setValues(currentSheet.pointValues); 

}
/**
 * regularPoints returns the number of points for a regular meeting, meeting point + points for SHPE werables
 * @param  {object} pastSheet [object that contains fields of the sheet we are comparing it to]
 * @param  {Number} i         [current row number]
 * @return {Number}           [Number of points that person will get]
 */
function regularPoints(pastSheet, i)
{
  var meetingPoint = 1;
  //regex for shirt
  var extraPts = extraPointsCheck(pastSheet.shirtColumnValues[i][0]);  
  return meetingPoint + extraPts;   
}

/**
 * extraPointsCheck checks if the value contains t-shirt or fleece and gives extra points
 * @param  {[String]} value [value containing if the person is wearing a shirt]
 * @return {[Integer]}      [number of extra points given]
 */
function extraPointsCheck(value)
{
  var extraPoints =0;
  var temp = String(value);
  if(temp.indexOf("SHPE T-Shirt") > -1)
    extraPoints++;
  if(temp.indexOf("SHPE Fleece") > -1)
    extraPoints++;
  return extraPoints;
}
/**
 * compareNames compares the last and first name and returns true if match
 * @param  {[object]} pastSheet    [custom object]
 * @param  {[int]} i               [index of past sheet loop]
 * @param  {[Object]} currentSheet [custom object]
 * @param  {[Int]} j               [index of current sheet loop]
 * @return {[bool]}                [boolean returning true or false]
 */
function compareNames(pastSheet,i,currentSheet,j)
{
    // '-' replace with space
    var pSfn = String(pastSheet.firstNameValues[i][0]).toLowerCase().trim().replace(/-/i, ' ');
    var pSln = String(pastSheet.lastNameValues[i][0]).toLowerCase().trim().replace(/-/i, ' ');
  
    var cSfn =  String(currentSheet.firstNameValues[j][0]).toLowerCase().trim();
    var cSln =String(currentSheet.lastNameValues[j][0]).toLowerCase().trim();
    return  pSfn ===  cSfn &&
           pSln === cSln;
}
/**
 * isMembershipComplete checks if tamuApplicant and nationalmember column are yes
 * @param  {[Object]} currentSheet [custom object]
 * @param  {[Int]} j               [index of current sheet loop]
 * @return {[bool]}                [boolean returning true or false]
 */
function isMembershipComplete(currentSheet,j)
{
  var tamuApplicant =  String(currentSheet.tamuApplicantValues[j][0]).toLowerCase();
  var nationalMember = String(currentSheet.nationalMemberValues[j][0]).toLowerCase();
  return tamuApplicant === "yes" && nationalMember === "yes";
}
function test ()
{
  var  inputData = {
            url:"https://docs.google.com/spreadsheets/d/1xrUn0cwDUe7PBSzBIs86V4sMveyszktpgrNtfM5MMHg/edit" ,
            pointsColumn: "E",
             currentSheet: {
              firstName: "B",
              lastName: "A",
              tamuApplicant: "C",
              nationalMember: "D" 
            },
            pastSheet: {
              firstName: "B",
              lastName: "C",
              shirt: "G",
              firstEvent: "E",
            }};
  compareSheet(inputData, regularPoints);
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Record Points').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
}
//takes in column letter and returns the value for that column
function letterToColumn(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

/*
            {
            url: ,
            pointsColumn: ,
             currentSheet: {
              firstName: ,
              lastName: ,
              tamuApplicant: ,
              nationalMember: 
            },
            pastSheet: {
              firstName: ,
              lastName: ,
              shirt: ,
              firstEvent: ,
            }
*/
function retrieveUserFields(data)
{
  //test();
  compareSheet(data, regularPoints);
  return true;
}


//will check if user has not payed all memberships  if so then send an email reminder to pay for it
function sendMembershipReminderEmails(data)
{
  //setup all roster values needed
  var FirstSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var currentSheet = sheet(FirstSheet, data.currentSheet.firstName ,3, data.currentSheet.lastName);
  currentSheet.pointValues = currentSheet.sheet.getSheetValues(currentSheet.firstNameRow, letterToColumn(targetColumn), 
   currentSheet.totalRows,1);
  setupRosterSheet(currentSheet, data.currentSheet.tamuApplicant, data.currentSheet.nationalMember);
  //go through all of the tamu and national member columns and check if any of them no or blank then we need 
  //to send an email 
  

}
//send an email 
function sendEmail(email, user)
{
    var subject = "";
    var body = "";
    // Send yourself an email with a link to the document.
    MailApp.sendEmail(email, subject, body);
}