''===========================================================
''Function to Create a Random Number with DateTime Stamp
''===========================================================
'Function fnRandomNumberWithDateTimeStamp()
'
''Find out the current date and time
'Dim sDate : sDate = Day(Now)
'Dim sMonth : sMonth = Month(Now)
'Dim sYear : sYear = Year(Now)
'Dim sHour : sHour = Hour(Now)
'Dim sMinute : sMinute = Minute(Now)
'Dim sSecond : sSecond = Second(Now)
'
''Create Random Number
'fnRandomNumberWithDateTimeStamp = Int(sDate & sMonth & sYear & sHour & sMinute & sSecond)
'
'End Function
''======================== End Function =====================
'
'Dim UserName, FirstName, LastName, Email, FullName



'Set Props = Browser("Browser").Page("SuccessFactors: Job Requisitio").WebTable("Job Requisition Summary").GetAllROProperties
'
'' Props contains the properties of the check box and their current values
'NumberOfProperties = Props.Count
'For i = 0 To NumberOfProperties - 1
'    Print Props(i).Name & ": " & Props(i).Value
'Next




'Browser("Browser").Page("SuccessFactors: Job Requisitio").WebTable("Job Requisition Summary").WaitProperty "visible",True, 3000


'AIUtil.FindTextBlock(DataTable.GlobalSheet.GetParameter("Categories")).Click
