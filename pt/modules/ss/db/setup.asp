<!--
Server-ML.co.uk PleaseTake System
Database Management Subsystem - Setup Commands
Copyright (c) Server-ML.co.uk 2006

setuptype(1) = Adding Departments
setuptype(2) = Adding Rooms
setuptype(3) = Adding Staff
setuptype(4) = Specifying Whether Or Not To Use Weekends
setuptype(5) = Specifying How Many Periods In A Day
setuptype(6) = Editing A Period
setuptype(7) = Determining Whether PINs Will Be Required Or Not
setuptype(8) = Determining Whether Full-Access Is Teaching Staff
setuptype(9) = Adding Full-Access Admin Details
setuptype(10) = Changing "Firstlogin" Value To Zero And Deleting Setup Account
-->

<!--#include virtual="/pt/modules/ss/p_s.inc"-->
<%
setuptype = request("setuptype")

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

if setuptype = "1" then

	FULL = request("full")

	if FULL = "" then
	response.redirect "/pt/admin/setup.asp?id=2&err=1"
	else
	end if

	if (request("short")) = "" then
	SHORT = FULL
	else
	SHORT = (request("short"))
	end if

	RSMAXNOSQL = "SELECT MAX(DEPTID) AS HIGHEST FROM Departments"
	Set RSMAXNO = Server.CreateObject("Adodb.RecordSet")
	RSMAXNO.Open RSMAXNOSQL, dataconn, adopenkeyset, adlockoptimistic

	if RSMAXNO("HIGHEST") <> "" then
		DEPTID = RSMAXNO("HIGHEST") + 1
	else
		DEPTID = 1
	end if

	addSQL = "INSERT INTO Departments ([SHORT],[FULL],[DEPTID]) Values ('" & SHORT & "','" & FULL & "','" & DEPTID & "')"
	dataconn.Execute(addSQL)

	RSMAXNO.close
	set RSMAXNO = nothing

	response.redirect "/pt/admin/setup.asp?id=2"

elseif setuptype = "2" then

	ROOM = request("room")

	if ROOM = "" then
	response.redirect "/pt/admin/setup.asp?id=3&err=1"
	else
	end if

	RSROOMSQL = "SELECT * FROM Rooms WHERE ROOMNO = '" & ROOM & "'"
	Set RSROOM = Server.CreateObject("Adodb.RecordSet")
	RSROOM.Open RSROOMSQL, dataconn, adopenkeyset, adlockoptimistic

	if RSROOM.RECORDCOUNT => 1 then
	response.redirect "/pt/admin/setup.asp?id=3&err=3"
	else
	end if

	addSQL = "INSERT INTO Rooms ([ROOMNO]) Values ('" & ROOM & "')"
	dataconn.Execute(addSQL)

	RSROOM.close
	set RSROOM = nothing

	response.redirect "/pt/admin/setup.asp?id=3"

elseif setuptype = "3" then

	TITLE = request("TITLE")
	FN = request("FN")
	LN = request("LN")
	DEPT = request("DEPT")
	CATEGORY = request("CATEGORY")
	DEFROOM = request("DEFROOM")
	ENTITLEMENT = request("ENTITLEMENT")
	
	RSCHECKSQL = "SELECT FN, LN FROM Timetables WHERE FN = '" & FN & "' AND LN = '" & LN & "'"
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
	
	if RSCHECK.RECORDCOUNT => 1 then
		RSCHECK.close
		set RSCHECK = nothing
		response.redirect "/pt/admin/setup.asp?id=4&err=1"
	else
		RSCHECK.close
		set RSCHECK = nothing
	end if

	if (FN = "") OR (LN = "") OR (CATEGORY = "") OR (DEPT = "") OR (TITLE = "") OR (ENTITLEMENT = "") THEN
		response.redirect "/pt/admin/setup.asp?id=4&err=2"
	else
	end if
	
	addSQL = "INSERT INTO Timetables ([TITLE],[FN],[LN],[DEPT],[CATEGORY],[DEFROOM],[ENTITLEMENT]) Values ('" & TITLE & "', '" & FN & "', '" & LN & "', '" & DEPT & "', '" & CATEGORY & "', '" & DEFROOM & "', '" & ENTITLEMENT & "')"
	dataconn.Execute(addSQL)

	RSCHECKSQL = "SELECT ID FROM Timetables WHERE FN = '" & FN & "' AND LN = '" & LN & "' AND CATEGORY = '" & CATEGORY & "' AND DEFROOM = '" & DEFROOM & "' AND DEPT = " & DEPT
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
	
	USER = RSCHECK("ID")
	
	RSCHECK.close
	set RSCHECK = nothing
	
	response.redirect "/pt/admin/setup.asp?id=4"

elseif setuptype = "4" then

	status = request("YN")

	Set update = settingsXML.documentElement.childNodes.item(0).childNodes.item(4)
	update.setAttribute "weekends", status

	settingsXML.Save(server.mappath("/pt/modules/xml/settings.xml"))

	response.redirect "/pt/admin/setup.asp?id=6"

elseif setuptype = "5" then

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

	response.redirect "/pt/admin/setup.asp?id=6&ok=1"

elseif setuptype = "6" then

	USER = request("USER")
	CLASSFIELD = request("PERIOD") & "_" & request("DAY")
	ROOMFIELD = "R" & request("PERIOD") & "_" & request("DAY")
	NAME = request("NAME")
	MODE = request("YN")

	if MODE = "1" then
		if NAME = "" then
			response.redirect "/pt/admin/setup.asp?id=7&part=3&err=1"
		else
		end if
	else
	end if
	
	if (request("YN")) = 1 then
		editSQL = "UPDATE Timetables SET [" & CLASSFIELD & "] = '" & NAME & "', [" & ROOMFIELD & "] = '" & request("ROOM") & "' WHERE ID = " & USER
		dataconn.Execute(editSQL)
	
		response.write "<script language='javascript' type='text/javascript'>opener.location.reload(true); self.close();</script>"
	else
		editSQL = "UPDATE Timetables SET [" & CLASSFIELD & "] = null, [" & ROOMFIELD & "] = null WHERE ID = " & USER
		dataconn.Execute(editSQL)
	
		response.write "<script language='javascript' type='text/javascript'>opener.location.reload(true); self.close();</script>"
	end if

elseif setuptype = "7" then

	PIN_new = request("YN")

	Set PIN = settingsXML.documentElement.childNodes.item(0).childNodes.item(4)
	PIN.setAttribute "pin", PIN_new

	settingsXML.Save(server.mappath("/pt/modules/xml/settings.xml"))

	response.redirect "/pt/admin/setup.asp?id=9"

elseif setuptype = "8" then

session("sess_fullaccess") = request("YN")

	if session("sess_fullaccess") = 1 then
		response.redirect "/pt/admin/setup.asp?id=9&part=1"
	else
		response.redirect "/pt/admin/setup.asp?id=9&part=2"
	end if

elseif setuptype = "9" then

	if (request("mode")) = "1" then

		if var_est_enabled_pin = 1 then

		UID = request("user")
		UN = request("UN")
		PW = request("PW")
		PWV = request("PWV")
		P1 = request("p1")
		P2 = request("p2")
		EM = request("EM")

		if (UN = "") OR (PW = "") OR (PWV = "") OR (P1 = "") OR (P2 = "") OR (EM = "") then
			response.redirect "/pt/admin/setup.asp?id=9&part=2&user=" & UID & "&err=1"
		else
			if (PW <> PWV) then
				response.redirect "/pt/admin/setup.asp?id=9&part=2&user=" & UID & "&err=2"
			else

				RSCHECKSQL = "SELECT * FROM Timetables WHERE ID = " & UID
				Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
				RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic

				if RSCHECK.RECORDCOUNT = 0 then
					RSCHECK.close
					set RSCHECK = nothing
					response.redirect "/pt/admin/setup.asp?id=9&part=2&user=" & UID & "&err=3"
				else

				RSCHECK2SQL = "SELECT * FROM Admin WHERE UN = '" & UN & "'"
				Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
				RSCHECK2.Open RSCHECK2SQL, userconn, adopenkeyset, adlockoptimistic

					if RSCHECK2.RECORDCOUNT => 1 then
						RSCHECK2.close
						set RSCHECK2 = nothing
						response.redirect "/pt/admin/setup.asp?id=9&part=2&user=" & UID & "&err=4"
					else
						addSQL = "INSERT INTO Admin ([TTID],[UN],[PW],[P1],[P2],[ACCLEVEL],[TITLE],[FN],[LN],[EMAIL],[DEPT]) Values ('" & RSCHECK("ID") & "','" & UN & "','" & PW & "','" & P1 & "','" & P2 & "','1','" & RSCHECK("TITLE") & "','" & RSCHECK("FN") & "','" & RSCHECK("LN") & "','" & EM & "@" & var_emaildomain1 & "','" & RSCHECK("DEPT") & "')"
						userconn.Execute(addSQL)
						RSCHECK.close
						set RSCHECK = nothing
						RSCHECK2.close
						set RSCHECK2 = nothing
						session("whereiam") = 10
						response.redirect "/pt/admin/setup.asp?id=10"
					end if
				end if
			end if
		end if

		else

		UID = request("user")
		UN = request("UN")
		PW = request("PW")
		PWV = request("PWV")
		EM = request("EM")

		if (UN = "") OR (PW = "") OR (PWV = "") OR (EM = "") then
			response.redirect "/pt/admin/setup.asp?id=9&part=2&user=" & UID & "&err=1"
		else
			if (PW <> PWV) then
				response.redirect "/pt/admin/setup.asp?id=9&part=2&user=" & UID & "&err=2"
			else

				RSCHECKSQL = "SELECT * FROM Timetables WHERE ID = " & UID
				Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
				RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic

				if RSCHECK.RECORDCOUNT = 0 then
					RSCHECK.close
					set RSCHECK = nothing
					response.redirect "/pt/admin/setup.asp?id=9&part=2&user=" & UID & "&err=3"
				else

				RSCHECK2SQL = "SELECT * FROM Admin WHERE UN = '" & UN & "'"
				Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
				RSCHECK2.Open RSCHECK2SQL, userconn, adopenkeyset, adlockoptimistic

					if RSCHECK2.RECORDCOUNT => 1 then
						RSCHECK2.close
						set RSCHECK2 = nothing
						response.redirect "/pt/admin/setup.asp?id=9&part=2&user=" & UID & "&err=4"
					else
						addSQL = "INSERT INTO Admin ([TTID],[UN],[PW],[ACCLEVEL],[TITLE],[FN],[LN],[EMAIL],[DEPT]) Values ('" & RSCHECK("ID") & "','" & UN & "','" & PW & "','1','" & RSCHECK("TITLE") & "','" & RSCHECK("FN") & "','" & RSCHECK("LN") & "','" & EM & "@" & var_emaildomain1 & "','" & RSCHECK("DEPT") & "')"
						userconn.Execute(addSQL)
						RSCHECK.close
						set RSCHECK = nothing
						RSCHECK2.close
						set RSCHECK2 = nothing
						session("whereiam") = 10
						response.redirect "/pt/admin/setup.asp?id=10"
					end if
				end if
			end if
		end if

		end if


	elseif (request("mode")) = "2" then

		if var_est_enabled_pin = 1 then

		TI = request("TITLE")
		FN = request("FN")
		LN = request("LN")
		UN = request("UN")
		PW = request("PW")
		PWV = request("PWV")
		P1 = request("p1")
		P2 = request("p2")
		EM = request("EM")

		if (FN = "") OR (LN = "") OR (UN = "") OR (PW = "") OR (PWV = "") OR (P1 = "") OR (P2 = "") OR (EM = "") then
			response.redirect "/pt/admin/setup.asp?id=9&part=2&err=1"
		else
			if (PW <> PWV) then
				response.redirect "/pt/admin/setup.asp?id=9&part=2&err=2"
			else

				RSCHECKSQL = "SELECT * FROM Admin WHERE UN = '" & UN & "'"
				Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
				RSCHECK.Open RSCHECKSQL, userconn, adopenkeyset, adlockoptimistic

					if RSCHECK.RECORDCOUNT => 1 then
						RSCHECK.close
						set RSCHECK = nothing
						response.redirect "/pt/admin/setup.asp?id=9&part=2&err=3"
					else
						addSQL = "INSERT INTO Admin ([UN],[PW],[P1],[P2],[ACCLEVEL],[TITLE],[FN],[LN],[EMAIL]) Values ('" & UN & "','" & PW & "','" & P1 & "','" & P2 & "','1','" & TI & "','" & FN & "','" & LN & "','" & EM & "@" & var_emaildomain1 & "')"
						userconn.Execute(addSQL)
						RSCHECK.close
						set RSCHECK = nothing
						session("whereiam") = 10
						response.redirect "/pt/admin/setup.asp?id=10"
					end if
			end if
		end if

		else

		TI = request("TITLE")
		FN = request("FN")
		LN = request("LN")
		UN = request("UN")
		PW = request("PW")
		PWV = request("PWV")
		EM = request("EM")

		if (FN = "") OR (LN = "") OR (UN = "") OR (PW = "") OR (PWV = "") OR (EM = "") then
			response.redirect "/pt/admin/setup.asp?id=9&part=2&err=1"
		else
			if (PW <> PWV) then
				response.redirect "/pt/admin/setup.asp?id=9&part=2&err=2"
			else

				RSCHECKSQL = "SELECT * FROM Admin WHERE UN = '" & UN & "'"
				Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
				RSCHECK.Open RSCHECKSQL, userconn, adopenkeyset, adlockoptimistic

					if RSCHECK.RECORDCOUNT => 1 then
						RSCHECK.close
						set RSCHECK = nothing
						response.redirect "/pt/admin/setup.asp?id=9&part=2&err=3"
					else
						addSQL = "INSERT INTO Admin ([UN],[PW],[ACCLEVEL],[TITLE],[FN],[LN],[EMAIL]) Values ('" & UN & "','" & PW & "','1','" & TI  & "','" & FN & "','" & LN & "','" & EM & "@" & var_emaildomain1 & "')"
						userconn.Execute(addSQL)
						RSCHECK.close
						set RSCHECK = nothing
						session("whereiam") = 10
						response.redirect "/pt/admin/setup.asp?id=10"
					end if
			end if
		end if

		end if



	end if

elseif setuptype = "10" then

	Set FLOGIN = settingsXML.documentElement.childNodes.item(0).childNodes.item(4)
	FLOGIN.setAttribute "firstlogin", "0"

	settingsXML.Save(server.mappath("/pt/modules/xml/settings.xml"))

	delSQL = "DELETE * FROM Admin WHERE UN = '" & settingsXML.documentElement.childNodes.item(0).childNodes.item(4).getAttribute("firstloginacc") & "'"
	userconn.Execute(delSQL)

	session("whereiam") = 11

	response.redirect "/pt/admin/setup.asp?id=11"

else

end if
%>