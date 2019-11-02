<!--
Server-ML.co.uk PleaseTake System
Database Management Subsystem - Editing
Copyright (c) Server-ML.co.uk 2006

edittype(1) = Editing Staff Details
edittype(2) = Editing Timetable Information
edittype(3) = Edting Attendance For Departmental Cover Wizard (Timetable)
edittype(4) = Editing XML Weekends Enabled/Disabled
edittype(5) = Editing How Many Periods In A Day
edittype(6) = Enabling/Disabling Features
edittype(7) = Changing Establishment Information
edittype(8) = Editing A Room's Name/Number
edittype(9) = Editing Account Details (Access Account)
edittype(10) = Editing A Department's Details
edittype(11) = Editing Outside Cover Availability
edittype(12) = Cover Summary - Changing Current Date
edittype(13) = Change My Settings (Admin)
edittype(14) = Edit Outside Cover Member's Details
edittype(15) = PleaseTake Slips - Changing Current Date
-->

<!--#include virtual="/pt/modules/ss/p_s.inc"-->
<%
edittype = request("edittype")

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

if edittype = "1" then
	USER = request("USER")
	TITLE = request("TITLE")
	FN = replace(request("FN"), "'", "''")
	LN = replace(request("LN"), "'", "''")
	DEPT = request("DEPT")
	CATEGORY = request("CATEGORY")
	DEFROOM = request("DEFROOM")
	ENTITLEMENT = request("ENTITLEMENT")
	
	editSQL = "UPDATE Timetables SET [TITLE] = '" & TITLE & "', [FN] = '" & FN & "', [LN] = '" & LN & "', [DEPT] = '" & DEPT & "', [CATEGORY] = '" & CATEGORY & "', [DEFROOM] = '" & DEFROOM & "', [ENTITLEMENT] = '" & ENTITLEMENT &"' WHERE ID = " & USER
	dataconn.Execute(editSQL)
	
	response.redirect "/pt/admin/staff.asp?id=5&good=1"

elseif edittype = "2" then
	USER = request("USER")
	CLASSFIELD = request("PERIOD") & "_" & request("DAY")
	ROOMFIELD = "R" & request("PERIOD") & "_" & request("DAY")
	NAME = replace(request("NAME"), "'", "''")
	MODE = request("YN")

	RSCHECKSQL = "SELECT * FROM Timetables WHERE ID = " & USER
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
	
	if (request("ROOM")) = "DEF" then
		ROOMSTR = RSCHECK("DEFROOM")
	elseif (request("ROOM")) = "NA" then
		ROOMSTR = "NA"
	else
		ROOMSTR = request("ROOM")
	end if

	if MODE = "1" then
		if NAME = "" then
			response.redirect "/pt/admin/popup.asp?id=1&err=1"
		else
		end if
	else
	end if
	
	if (request("YN")) = 1 then
		editSQL = "UPDATE Timetables SET [" & CLASSFIELD & "] = '" & NAME & "', [" & ROOMFIELD & "] = '" & ROOMSTR & "' WHERE ID = " & USER
		dataconn.Execute(editSQL)
	
		response.write "<script language='javascript' type='text/javascript'>opener.location.reload(true); self.close();</script>"
	else
		editSQL = "UPDATE Timetables SET [" & CLASSFIELD & "] = null, [" & ROOMFIELD & "] = null WHERE ID = " & USER
		dataconn.Execute(editSQL)
	
		response.write "<script language='javascript' type='text/javascript'>opener.location.reload(true); self.close();</script>"
	end if
	
	RSCHECK.close
	set RSCHECK = nothing

elseif edittype = "3" then

	USER = request("user")
	
		if (request("whole")) <> "1" then
			PERIOD = request("period")
			DAYDOW = request("day")
			DAYDATE = request("coverday")
			FIELDNAME = PERIOD & "_" & DAYDOW
			RSCHECKSQL = "SELECT * FROM Attendance WHERE USER = " & USER & " AND DAY = " & DAYDOW & " AND DAYDATE = #" & SQLDate(DAYDATE) & "#"
			Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
			RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
	
			if RSCHECK(PERIOD & "_" & DAYDOW) <> "" then
				editSQL = "UPDATE Attendance SET [" & FIELDNAME & "] = null WHERE USER = " & USER & " AND DAY = " & DAYDOW & " AND DAYDATE = #" & SQLDate(DAYDATE) & "#"
				dataconn.Execute(editSQL)
				delSQL = "DELETE * FROM Cover WHERE DAY = " & DAYDOW & " AND DAYDATE = #" & SQLDate(DAYDATE) & "# AND PERIOD = " & PERIOD & " AND FOR = " & USER
				dataconn.Execute(delSQL)
			else
				editSQL = "UPDATE Attendance SET [" & FIELDNAME & "] = 'A' WHERE USER = " & USER & " AND DAY = " & DAYDOW & " AND DAYDATE = #" & SQLDate(DAYDATE) & "#"
				dataconn.Execute(editSQL)
			end if

			RSCHECK.close
			set RSCHECK = nothing
	
			if (request("type")) = "2" then
				response.redirect "/pt/admin/cover.asp?id=3&type=2&coverday=" & DAYDATE & "&dow=" & DAYDOW
			else
				response.redirect "/pt/admin/cover.asp?id=3&type=1&coverday=" & DAYDATE & "&dow=" & DAYDOW
			end if
		else
			TOTAL = request("total")
			for i=1 to TOTAL
				DAYDOW = request("day")
				DAYDATE = request("coverday")
				FIELDNAME = i & "_" & DAYDOW
				
				editSQL = "UPDATE Attendance SET [" & FIELDNAME & "] = 'A' WHERE USER = " & USER & " AND DAY = " & DAYDOW & " AND DAYDATE = #" & SQLDate(DAYDATE) & "#"
				dataconn.Execute(editSQL)

			next
			
			if (request("type")) = "2" then
				response.redirect "/pt/admin/cover.asp?id=3&type=2&coverday=" & DAYDATE & "&dow=" & DAYDOW
			else
				response.redirect "/pt/admin/cover.asp?id=3&type=1&coverday=" & DAYDATE & "&dow=" & DAYDOW
			end if
		end if

elseif edittype = "4" then

	status = request("YN")

	Set update = settingsXML.documentElement.childNodes.item(0).childNodes.item(4)
	update.setAttribute "weekends", status

	settingsXML.Save(server.mappath("/pt/modules/xml/settings.xml"))

	response.redirect "/pt/admin/settings.asp?id=4&ok=1"

elseif edittype = "5" then

	if (request("period")) = "" then

		RSCHECKSQL = "SELECT * FROM Periods WHERE DAYID = " & request("DOW")
		Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
		RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic

		if RSCHECK("TOTALS") = "0" then

			editSQL = "UPDATE Periods SET [Totals] = '" & var_maxperiods & "' WHERE DAYID = " & request("DOW")
			dataconn.Execute(editSQL)

		else

			editSQL = "UPDATE Periods SET [Totals] = '0' WHERE DAYID = " & request("DOW")
			dataconn.Execute(editSQL)

		end if

		RSCHECK.close
		set RSCHECK = nothing

	elseif (request("DOW")) = "" then

		if var_est_enabled_weekends = "1" then

			editSQL = "UPDATE Periods SET [Totals] = " & request("period")
			dataconn.Execute(editSQL)

		else

			for i = 2 to 6

				editSQL = "UPDATE Periods SET [Totals] = " & request("period") & " WHERE DAYID = " & i
				dataconn.Execute(editSQL)

			next

		end if

	else

		editSQL = "UPDATE Periods SET [Totals] = '" & request("period") & "' WHERE DAYID = " & request("DOW")
		dataconn.Execute(editSQL)

	end if

	response.redirect "/pt/admin/settings.asp?id=4&ok=1"

elseif edittype = "6" then

	entire_new = request("SYSTEM")
	signup_new = request("SIGNUP")
	PIN_new = request("PIN")

	Set entire = settingsXML.documentElement.childNodes.item(0).childNodes.item(4)
	entire.setAttribute "entire", entire_new

	Set signup = settingsXML.documentElement.childNodes.item(0).childNodes.item(4)
	signup.setAttribute "signup", signup_new

	Set PIN = settingsXML.documentElement.childNodes.item(0).childNodes.item(4)
	PIN.setAttribute "pin", PIN_new

	settingsXML.Save(server.mappath("/pt/modules/xml/settings.xml"))

	response.redirect "/pt/admin/settings.asp?id=3&ok=1"

elseif edittype = "7" then

	full_new = request("FULL")
	short_new = request("SHORT")
	std_new = request("STD")
	admin_new = request("ADMIN")

	set full = settingsXML.documentElement.childNodes.item(0).childNodes.item(3)
	full.text = full_new

	Set short = settingsXML.documentElement.childNodes.item(0).childNodes.item(3)
	short.setAttribute "short", short_new

	Set std = settingsXML.documentElement.childNodes.item(0).childNodes.item(6)
	std.setAttribute "std", std_new

	Set admin = settingsXML.documentElement.childNodes.item(0).childNodes.item(6)
	admin.setAttribute "admin", admin_new

	settingsXML.Save(server.mappath("/pt/modules/xml/settings.xml"))

	response.redirect "/pt/admin/settings.asp?id=5&ok=1"

elseif edittype = "8" then

	ROOM = replace(request("ROOM"), "'", "''")
	ROOMID = replace(request("ROOMID"), "'", "''")

	if ROOM = "" then
		response.redirect "/pt/admin/rooms.asp?id=2&room=" & ROOMID & "&err=1"
	else
		editSQL = "UPDATE Rooms SET [ROOMNO] = '" & ROOM & "' WHERE ID = " & ROOMID
		dataconn.Execute(editSQL)
		response.redirect "/pt/admin/rooms.asp?id=2&gd=2"
	end if

elseif edittype = "9" then

	'i don't know how to do this bit well

elseif edittype = "10" then

	'EDIT A ROOM

elseif edittype = "11" then

	staffid = request("uid")
	changetype = request("type")
	currfield = "D_" & DOW

		if (changetype = "1a") then
			editSQL = "UPDATE OCover SET [ENTIRE] = 0, [" & currfield & "] = '0' WHERE ID = " & staffid
			dataconn.Execute(editSQL)
			response.redirect "/pt/admin/ocover.asp?id=1"
		elseif (changetype = "1b") then
			edit1SQL = "UPDATE OCover SET [ENTIRE] = 0, [D_1] = 0, [D_2] = 0, [D_3] = 0, [D_4] = 0, [D_5] = 0, [D_6] = 0, [D_7] = 0 WHERE ID = " & staffid
			edit2SQL = "UPDATE OCover SET [ENTIRE] = 0, [" & currfield & "] = '1' WHERE ID = " & staffid
			dataconn.Execute(edit1SQL)
			dataconn.Execute(edit2SQL)
			response.redirect "/pt/admin/ocover.asp?id=1"
		elseif (changetype = "2a") then
			editSQL = "UPDATE OCover SET [ENTIRE] = 0, [" & currfield & "] = '0' WHERE ID = " & staffid
			dataconn.Execute(editSQL)
			response.redirect "/pt/admin/ocover.asp?id=1"
		elseif (changetype = "2b") then
			editSQL = "UPDATE OCover SET [ENTIRE] = 1, [D_1] = 1, [D_2] = 1, [D_3] = 1, [D_4] = 1, [D_5] = 1, [D_6] = 1, [D_7] = 1 WHERE ID = " & staffid
			dataconn.Execute(editSQL)
			response.redirect "/pt/admin/ocover.asp?id=1"
		elseif (changetype = "3a") then
			editSQL = "UPDATE OCover SET [" & currfield & "] = '1' WHERE ID = " & staffid
			dataconn.Execute(editSQL)
			response.redirect "/pt/admin/ocover.asp?id=1"
		elseif (changetype = "3b") then
			editSQL = "UPDATE OCover SET [ENTIRE] = 0, [D_1] = 0, [D_2] = 0, [D_3] = 0, [D_4] = 0, [D_5] = 0, [D_6] = 0, [D_7] = 0 WHERE ID = " & staffid
			dataconn.Execute(editSQL)
			response.redirect "/pt/admin/ocover.asp?id=1"
		elseif (changetype = "4a") then
			editSQL = "UPDATE OCover SET [" & currfield & "] = '1' WHERE ID = " & staffid
			dataconn.Execute(editSQL)
			response.redirect "/pt/admin/ocover.asp?id=1"
		elseif (changetype= "4b") then
			editSQL = "UPDATE OCover SET [ENTIRE] = 1, [D_1] = 1, [D_2] = 1, [D_3] = 1, [D_4] = 1, [D_5] = 1, [D_6] = 1, [D_7] = 1 WHERE ID = " & staffid
			dataconn.Execute(editSQL)
			response.redirect "/pt/admin/ocover.asp?id=1"
		else
			response.redirect "/pt/admin/ocover.asp?id=1"
		end if

elseif edittype = "12" then

	if (request("REP3_CHOICE") = "") then
		response.write "<script language='javascript' type='text/javascript'>history.back();</script>"
	else
		dow = left(request("REP3_CHOICE"),1)
		daydate = mid(request("REP3_CHOICE"),3)

		response.redirect "/pt/admin/reports.asp?id=3&dow=" & dow & "&daydate=" & daydate
	end if

elseif edittype = "13" then

if session("sess_ttid") <> undefined then
	if var_est_enabled_pin <> "1" then

	RSUSERSQL = "SELECT * FROM Timetables WHERE ID = " & session("sess_ttid")
	RSACCSQL = "SELECT * FROM Admin WHERE UN = '" & session("sess_un") & "'"

	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic

	Set RSACC = Server.CreateObject("Adodb.RecordSet")
	RSACC.Open RSACCSQL, userconn, adopenkeyset, adlockoptimistic

		if RSUSER.RECORDCOUNT = 0 then
			response.redirect "/pt/admin/settings.asp?id=2&err=4"
		else
			var_title = request("title")
			var_fn = request("fn")
			var_ln = request("ln")
			var_email = request("email")
			var_dept = request("dept")
			var_pos = request("category")
			var_defroom = request("defroom")

			var_passo = request("password_o")
			var_passn = request("password_n")
			var_passc = request("password_c")
			
			if (var_fn = undefined) or (var_ln = undefined) or (var_email = undefined) then
				response.redirect "/pt/admin/settings.asp?id=2&err=1"
			else
				var_email = var_email & "@" & var_emaildomain1
				if var_passo = undefined then
					editSQL = "UPDATE Admin SET [Title] = '" & var_title & "', [FN] = '" & var_fn & "', [LN] = '" & var_ln & "', [EMAIL] = '" & var_email & "', [DEPT] = '" & var_dept & "' WHERE UN = '" & session("sess_un") & "'"
					userconn.execute(editSQL)
					editSQL = "UPDATE Timetables SET [Title] = '" & var_title & "', [FN] = '" & var_fn & "', [LN] = '" & var_ln & "', [DEPT] = '" & var_dept & "', [CATEGORY] = '" & var_pos & "', [DEFROOM] = '" & var_defroom & "' WHERE ID = " & session("sess_ttid")
					dataconn.execute(editSQL)
				
					response.redirect "/pt/admin/settings.asp?id=2&gd=1"
				else
					if (var_passn = undefined) or (var_passc = undefined) then
						response.redirect "/pt/admin/settings.asp?id=2&err=1"
					else
						if (var_passn <> var_passc) then
							response.redirect "/pt/admin/settings.asp?id=2&err=2"
						else
							if RSACC("PW") <> var_passo then
								response.redirect "/pt/admin/settings.asp?id=2&err=3"
							else
								editSQL = "UPDATE Admin SET [PW] = '" & var_passn & "', [Title] = '" & var_title & "', [FN] = '" & var_fn & "', [LN] = '" & var_ln & "', [EMAIL] = '" & var_email & "', [DEPT] = '" & var_dept & "' WHERE UN = '" & session("sess_un") & "'"
								userconn.execute(editSQL)
								editSQL = "UPDATE Timetables SET [Title] = '" & var_title & "', [FN] = '" & var_fn & "', [LN] = '" & var_ln & "', [DEPT] = '" & var_dept & "', [CATEGORY] = '" & var_pos & "', [DEFROOM] = '" & var_defroom & "' WHERE ID = " & session("sess_ttid")
								dataconn.execute(editSQL)
								
								response.redirect "/pt/admin/settings.asp?id=2&gd=1"
							end if
						end if
					end if
				end if
			end if
		end if
	else
%>
teach, PIn
<%
	end if
else
%>
noteach
<%
end if


elseif edittype = "14" then
	if (request("uid") = "") then
		response.redirect "/pt/admin/ocover.asp?id=1&err=3"
	else
		if (isNumeric(request("uid")) = False) then
			response.redirect "/pt/admin/ocover.asp?id=1&err=3"
		else
			RSCHECKSQL = "SELECT * FROM OCover WHERE ID = " & request("uid")
			Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
			RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic

			if RSCHECK.RECORDCOUNT = 0 then
				RSCHECK.close
				set RSCHECK = nothing
				response.redirect "/pt/admin/ocover.asp?id=1&err=2"
			else
				if (request("FN") = "") OR (request("LN") = "") OR (request("ENTITLEMENT") = "") then
					response.redirect "/pt/admin/ocover.asp?id=2&uid=" & request("uid") & "&err=1"
				else
					FN = replace(request("FN"), "'", "''")
					LN = replace(request("LN"), "'", "''")
				
					editSQL = "UPDATE OCover SET [TITLE] = '" & request("TITLE") & "', [FN] = '" & FN & "', [LN] = '" & LN & "', [ENTITLEMENT] = '" & request("ENTITLEMENT") & "' WHERE ID = " & request("uid")
					dataconn.execute(editSQL)
					
					response.redirect "/pt/admin/ocover.asp?id=1&ok=1"
				end if
			
				RSCHECK.close
				set RSCHECK = nothing
			end if
		end if
	end if

elseif edittype = "15" then

	if (request("REP3_CHOICE") = "") then
		response.write "<script language='javascript' type='text/javascript'>history.back();</script>"
	else
		dow = left(request("REP3_CHOICE"),1)
		daydate = mid(request("REP3_CHOICE"),3)

		response.redirect "/pt/admin/reports.asp?id=2&dow=" & dow & "&daydate=" & daydate
	end if

else

end if
%>