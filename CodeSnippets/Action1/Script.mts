'===========================================================
'Function for creating a number at run time based on current time down to the second, to allow for a unique number each time the script is run
'===========================================================
Function fnRandomNumberWithDateTimeStamp()

'Find out the current date and time
Dim sDate : sDate = Day(Now)
Dim sMonth : sMonth = Month(Now)
Dim sYear : sYear = Year(Now)
Dim sHour : sHour = Hour(Now)
Dim sMinute : sMinute = Minute(Now)
Dim sSecond : sSecond = Second(Now)

'Create Random Number
fnRandomNumberWithDateTimeStamp = Int(sDate & sMonth & sYear & sHour & sMinute & sSecond)

'======================== End Function =====================
End Function

'===========================================================
'Function for debugging properties at run time to output to the log, once script is working, can be deleted
'===========================================================
Function PropertiesDebug

'Dim CPActual, CPExpected
'Debug code to determine why the checkpoint was failing, turns out that there is a trailing space in the application code that the result HTML was trimming when displaying expected vs. actual
'CPExpected = "'" & DataTable.Value ("FullName") & "'"										'Set the variable for what is in the data table, enclose with single quotes so we can find leading/trailing spaces
'CPActual = Browser("Browser").Page("SuccessFactors: Candidates").Link("CandidateName").GetROProperty("text")	'Get the actual text from the object at run time
'CPActual = "'" & CPActual & "'"															'Set the variable for what is the object property at run time enclosed with a single quotes so we can find leading/trailing spaces
'Print "Expected is " & CPExpected															'Output the expected value to the output log
'Print "Actual is " & CPActual																'Output the actual value to the output log
	
End Function

'===========================================================
'Function for closing all open browsers, then launching the browser so that we are always sure we're in the right startup state for the browser
'===========================================================
Function AppContext

While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
SystemUtil.Run "CHROME.exe" ,"","","",3														'launch Chrome, could be data drive to launch other browser (e.g. Firefox)
Set AppContext=Browser("CreationTime:=0")													'Set the variable for what application (in this case the browser) we are acting upon

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate "about:blank"															'Navigate to the application URL
AppContext.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																				'Wait for the browser to stop spinning

End Function

