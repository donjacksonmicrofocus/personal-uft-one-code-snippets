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
AIUtil.SetContext AppContext																'Tell the AI engine to point at the application

End Function


Set clipBoard = CreateObject("Mercury.Clipboard") 


Public Function OpenPDFDocument(ByRef pdfPath)
      SystemUtil.Run pdfPath
End Function

Public Function GoToHome
    Window("Adobe Acrobat Pro DC").Type micHome
	wait 0,100
End Function

Public Function FindCheckNumber(ByRef text)
	checkNumberTextIndex = InStr(text, "Check No")
	If checkNumberTextIndex <> 0 Then
		checkNumberIndex = checkNumberTextIndex + 10
		checkNumber = Mid(text, checkNumberIndex, 10)
		if IsNUmeric(checkNumber) then
			FindCheckNumber = CLng(checkNumber)
		Else
			Reporter.ReportEvent micFail, "FindCheckNumber", "FindCheckNumber was not able to read check number as a number. The read value is " & checkNumber  & " and the search text is:" & vblf & text 
			FindCheckNumber = 0
		End  If
	Else
		FindCheckNumber = 0
		Reporter.ReportEvent micFail, "FindCheckNumber", "FindCheckNumber was not able to locate the 'Check No' field. The search text is:" & vblf & text 
	End If
End Function

Public Function GetTextFromCurrentPage
    Window("Adobe Acrobat Pro DC").WinObject("AVPageView").Type micCtrlDwn + "a" + micCtrlUp
    Window("Adobe Acrobat Pro DC").WinObject("AVPageView").Type micCtrlDwn + "c" + micCtrlUp
    textOfCurrentPage = clipBoard.GetText
    Window("Adobe Acrobat Pro DC").WinObject("AVPageView").Type micShiftDwn + micCtrlDwn + "a" + micShiftUp + micCtrlUp
    clipBoard.Clear
    GetTextFromCurrentPage = textOfCurrentPage
End Function


Public Function FindCheckNumberFromCurrentPage
     currentPageText = GetTextFromCurrentPage()
     FindCheckNumberFromCurrentPage = FindCheckNumber(currentPageText)
End Function

Public Function GetCurrentPageNumber
     currentPageNumber = Window("Adobe Acrobat Pro DC").WinEdit("PageNumber").GetROProperty("text")
     GetCurrentPageNumber = CInt(currentPageNumber)
End Function

Public Function SetCurrentPageNumber(ByRef pageNumber)
     Window("Adobe Acrobat Pro DC").WinEdit("PageNumber").Set pageNumber
     Window("Adobe Acrobat Pro DC").WinEdit("PageNumber").Type micReturn
     wait 0, 200
     if Window("Adobe Acrobat Pro DC").Dialog("Adobe Acrobat").WinButton("OK").Exist(0) then
     	Window("Adobe Acrobat Pro DC").Dialog("Adobe Acrobat").WinButton("OK").Click
     	SetCurrentPageNumber = false
     Else
     	SetCurrentPageNumber = true
     End  If

End Function

Public Function FindPageForTextContent(ByRef text)
	  GoToHome()
      Window("Adobe Acrobat Pro DC").Type micCtrlDwn + "f" + micCtrlUp 'find tool
      wait 0, 100
      Window("Adobe Acrobat Pro DC").WinEdit("FindText").Set text 'enter text to find
      Window("Adobe Acrobat Pro DC").WinEdit("FindText").Type micReturn 'start searching
      wait 0,200
      FindPageForTextContent = GetCurrentPageNumber()
      Window("Adobe Acrobat Pro DC").Type micEsc  'hide find tool
End Function

Public Function IsSignatureValidForCurrentPage
	Window("Adobe Acrobat Pro DC").WinObject("AVPageView").Type micPgDwn
	if Window("Adobe Acrobat Pro DC").InsightObject("Signature1").Exist(0) then
       	IsSignatureValidForCurrentPage = true
   	Else
       	IsSignatureValidForCurrentPage = false
       	Reporter.ReportEvent micFail, "IsSignatureValidForCurrentPage", "Signature does not match"
    End  If
End Function

Public Function PrepareForReplay
       Window("Adobe Acrobat Pro DC").Maximize
       wait 0,100
       Window("Adobe Acrobat Pro DC").Type micAltDwn + "v" + "p" + "s" + micAltUp 'single page view
       wait 0, 100
       Window("Adobe Acrobat Pro DC").Type micCtrlDwn + "1" + micCtrlUp 'zoom to actual size
       wait 0,100
End Function
 
Public Function VerifyCheckNumberAndSignatureFromPage(ByRef pageNumber, ByRef checkNumber)
       pageIsValid = SetCurrentPageNumber(pageNumber)
       If pageIsValid Then
       		currentCheckNumber = FindCheckNumberFromCurrentPage()
       		If currentCheckNumber <> 0 Then
       			If currentCheckNumber = checkNumber Then
       				If IsSignatureValidForCurrentPage() Then
       					Reporter.ReportEvent micPass, "CheckForNumberAndSignature", "Signature and check number match!"
       				Else
       					Reporter.ReportEvent micFail, "CheckForNumberAndSignature", "Signature does not match!"
       				End If
       			Else
       				Reporter.ReportEvent micFail, "CheckForNumberAndSignature", "Check number does not match. Current:" &  currentCheckNumber & " Expected:" & checkNumber
       			End If
       		Else
       			Reporter.ReportEvent micFail, "CheckForNumberAndSignature", "Cannot extract the check number from the current page"
       		End If
       Else
       		Reporter.ReportEvent micFail, "CheckForNumberAndSignature", "Invalid page number"
       End If
End Function
Function RunThePDFVerificationProcess

OpenPDFDocument "C:\Check.pdf"
PrepareForReplay 'run this first time to ensure right settings in adobe pdf

print FindPageForTextContent("1000013580") ' call this if you need to find a particular page number for specific text
VerifyCheckNumberAndSignatureFromPage 3,1000013571 ' call this to verify that a particular pdf page maches a check number and a signature

End Function 

