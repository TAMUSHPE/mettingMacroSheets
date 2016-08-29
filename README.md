# mettingMacroSheets
Google sheets Macro used for automatically adding points to membership sheet from meetings

##Goal
Create a macro that is able to automatically populate points from the  meetings to the users point sheet.

#Description
Script is added to main google sheet as a script with the sidebar.html file.   
It will then add a custom Menu in which you hover over and has "Points meeting macro"
Which will open a side bar with fields to enter:  
"Url of target Meeting Sheet" aka the url of the sheet of the meeting  
"Column for Points"  aka the column where you would want the points to be 

There is seetings for each sheet to change the defaults if you click the buttons and enter different columns.

<b>if you have a file without a first event column</b> then change the deafult to <b>NONE</b> and it won't check that column but still do points     
You can also just put a column with nothing on it and it will still work   

Key:
After runining the script it might highlight some rows in the meeting sheet.  
If a row is highlighted yellow: 
* It means that you should make sure that this user isn't a member since they might have typed their name wrong   

If a row is highlighted red: 
* It means the user said that was there first meeting therefore he probably isn't a member

<img src='https://github.com/TAMUSHPE/mettingMacroSheets/blob/master/pointsMacro2.gif' title='Video Walkthrough' width='' alt='Video Walkthrough' />
GIF created with [LiceCap](http://www.cockos.com/licecap/).

##References
[Google Sheets API](https://developers.google.com/apps-script/guides/sheets)
