<%
current = now()

date_full = dateAdd("d", i, current)
date_day = datePart("d", current)
date_month = datePart("m", current)
date_year = datePart("yyyy", current)
date_hour = datePart("h", current)

Function SQLDate(StartDate)
	Dim strSeparator
	Dim strTheMonth
	Dim strTheYear
	Dim strTheDay
	Dim strFinalDate
	strSeparator = "/"
	strTheMonth = Month(StartDate)
	strTheYear = Year(StartDate)
	strTheDay = Day(StartDate)
	strFinalDate = strTheMonth & strSeparator & strTheDay & strSeparator & strTheYear
	SQLDate = strFinalDate
End Function

Function Cal_GetMonth(CurrentDate)
	Dim strMonth
		Select Case CurrentDate
		Case "1"
			strMonth = "January"
		Case "2"
			strMonth = "Feburary"
		Case "3"
			strMonth = "March"
		Case "4"
			strMonth = "April"
		Case "5"
			strMonth = "May"
		Case "6"
			strMonth = "June"
		Case "7"
			strMonth = "July"
		Case "8"
			strMonth = "August"
		Case "9"
			strMonth = "September"
		Case "10"
			strMonth = "October"
		Case "11"
			strMonth = "November"
		Case "12"
			strMonth = "December"
		End Select
	Cal_GetMonth = strMonth
End Function

Function GetMonth(CurrentDate)
	Dim strMonth
		Select Case Month(CurrentDate)
		Case "1"
			strMonth = "January"
		Case "2"
			strMonth = "Feburary"
		Case "3"
			strMonth = "March"
		Case "4"
			strMonth = "April"
		Case "5"
			strMonth = "May"
		Case "6"
			strMonth = "June"
		Case "7"
			strMonth = "July"
		Case "8"
			strMonth = "August"
		Case "9"
			strMonth = "September"
		Case "10"
			strMonth = "October"
		Case "11"
			strMonth = "November"
		Case "12"
			strMonth = "December"
		End Select
	GetMonth = strMonth
End Function

Function GetDOW(CurrentDate)
	Dim strDOW
		Select Case WeekDay(CurrentDate)
		Case "1"
			strDOW = "Sunday"
		Case "2"
			strDOW = "Monday"
		Case "3"
			strDOW = "Tuesday"
		Case "4"
			strDOW = "Wednesday"
		Case "5"
			strDOW = "Thursday"
		Case "6"
			strDOW = "Friday"
		Case "7"
			strDOW = "Saturday"
		End Select
	GetDOW = strDOW
End Function

if GETDOW(date()) = "Sunday" then
	DOW = 1
elseif GETDOW(date()) = "Monday" then
	DOW = 2
elseif GETDOW(date()) = "Tuesday" then
	DOW = 3
elseif GETDOW(date()) = "Wednesday" then
	DOW = 4
elseif GETDOW(date()) = "Thursday" then
	DOW = 5
elseif GETDOW(date()) = "Friday" then
	DOW = 6
elseif GETDOW(date()) = "Saturday" then
	DOW = 7
end if

if GETDOW(date()) = "Sunday" then
	DOW_SUM = 1
elseif GETDOW(date()) = "Monday" then
	DOW_SUM = 2
elseif GETDOW(date()) = "Tuesday" then
	DOW_SUM = 3
elseif GETDOW(date()) = "Wednesday" then
	DOW_SUM = 4
elseif GETDOW(date()) = "Thursday" then
	DOW_SUM = 5
elseif GETDOW(date()) = "Friday" then
	DOW_SUM = 6
elseif GETDOW(date()) = "Saturday" then
	DOW_SUM = 7
end if

Function GetDOWNo(FuncDate)
		Select Case WeekDay(FuncDate)
		Case "1"
			DOW = "1"
		Case "2"
			DOW = "2"
		Case "3"
			DOW = "3"
		Case "4"
			DOW = "4"
		Case "5"
			DOW = "5"
		Case "6"
			DOW = "6"
		Case "7"
			DOW = "7"
		End Select
	GetDOWNo = DOW
End Function

Function GetWeekNo(FuncDate)
	weekno = datepart("ww",FuncDate)
	GetWeekNo = weekno
End Function

Function GetWeekStart(FuncDate)
Select Case getDOWNo(FuncDate)
	Case "1"

		startdate = FuncDate

	Case "2"
	
		startdate = DateAdd("d",-1,FuncDate)

	Case "3"

		startdate = DateAdd("d",-2,FuncDate)

	Case "4"

		startdate = DateAdd("d",-3,FuncDate)

	Case "5"

		startdate = DateAdd("d",-4,FuncDate)

	Case "6"

		startdate = DateAdd("d",-5,FuncDate)

	Case "7"

		startdate = DateAdd("d",-6,FuncDate)

	Case Else

End Select

	GetWeekStart = startdate
End Function

Function GetWeekEnd(FuncDate)
Select Case getDOWNo(FuncDate)
	Case "1"

		enddate = DateAdd("d",+6,FuncDate)

	Case "2"
	
		enddate = DateAdd("d",+5,FuncDate)

	Case "3"

		enddate = DateAdd("d",+4,FuncDate)

	Case "4"

		enddate = DateAdd("d",+3,FuncDate)

	Case "5"

		enddate = DateAdd("d",+2,FuncDate)

	Case "6"

		enddate = DateAdd("d",+1,FuncDate)

	Case "7"

		enddate = FuncDate

	Case Else

End Select

	GetWeekEnd = enddate
End Function

dbxml_date = date()
dbxml_time = FormatDateTime(current,4)

date_without0 = date_day & "/" & date_month & "/" & date_year

display_date = GetDOW(date_full) & ", "& GetMonth(date_full) & "&nbsp;" & date_day & "&nbsp;" & date_year
display_time = dbxml_time & " " & var_tz
%>