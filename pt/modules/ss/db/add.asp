<!--
Server-ML.co.uk PleaseTake System
Database Management Subsystem - Adding
Copyright (c) Server-ML.co.uk 2006

addtype(1) = Adding Staff Details
addtype(2) = Department Cover Wizard - Selecting Who Is Absent
addtype(3) = Adding Cover Staff Into Cover Table
addtype(4) = Add A Room
addtype(5) = Add From Signup Wizard (Account Details)
addtype(6) = Add An Account (Create Account From Staff Details)
addtype(7) = Add A Department
addtype(8) = Add An Outside Cover Member
-->

<!--#include virtual="/pt/modules/ss/p_s.inc"-->
<%
addtype = request("addtype")

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

if addtype = "1" then

	TITLE = request("TITLE")
	FN = request("FN")
	LN = request("LN")
	DEPT = request("DEPT")
	CATEGORY = request("CATEGORY")
	DEFROOM = request("DEFROOM")
	ENTITLEMENT = request("ENTITLEMENT")
	CREATE = request("CREATE")
	UN = request("UN")
	PW = request("PW")
	P1 = request("P1")
	P2 = request("P2")
	EMAIL = request("EMAIL") & "@" & var_emaildomain1
	ADDTYPE = request("TYPE")

	if CREATE = "1" then
		if var_est_enabled_pin = "1" then

		if (FN = "") OR (LN = "") OR (ENTITLEMENT = "") OR (UN = "") OR (PW = "") OR (P1 = "") OR (P2 = "") THEN
			response.redirect "/pt/admin/staff.asp?id=6&err=2"
		else
		end if

		RSCHECKSQL = "SELECT FN, LN FROM Timetables WHERE FN = '" & FN & "' AND LN = '" & LN & "'"
		Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
		RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
		
		if (ADDTYPE = "2") OR (ADDTYPE = "3") then
			RSCHECK2SQL = "SELECT UN FROM Admin WHERE UN = '" & UN & "'"
		else
			RSCHECK2SQL = "SELECT UN FROM Users WHERE UN = '" & UN & "'"
		end if
		
		Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
		RSCHECK2.Open RSCHECK2SQL, userconn, adopenkeyset, adlockoptimistic
		
		if RSCHECK.RECORDCOUNT => 1 then
			RSCHECK.close
			set RSCHECK = nothing
			response.redirect "/pt/admin/staff.asp?id=6&err=1"
		elseif RSCHECK2.RECORDCOUNT => 1 then
			RSCHECK.close
			set RSCHECK = nothing
			RSCHECK2.close
			set RSCHECK2 = nothing
			response.redirect "/pt/admin/staff.asp?id=6&err=1"
		else
			RSCHECK.close
			set RSCHECK = nothing
		end if
		
		addSQL = "INSERT INTO Timetables ([TITLE],[FN],[LN],[DEPT],[CATEGORY],[DEFROOM],[ENTITLEMENT]) Values ('" & TITLE & "', '" & FN & "', '" & LN & "', '" & DEPT & "', '" & CATEGORY & "', '" & DEFROOM & "', '" & ENTITLEMENT & "')"
		dataconn.Execute(addSQL)
	
		RSCHECKSQL = "SELECT ID FROM Timetables WHERE FN = '" & FN & "' AND LN = '" & LN & "' AND CATEGORY = '" & CATEGORY & "' AND DEFROOM = '" & DEFROOM & "' AND DEPT = " & DEPT
		Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
		RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic

			if ADDTYPE = "2" then
				adminlevel = 2
			elseif ADDTYPE = "3" then
				adminlevel = 1
			end if

		if (ADDTYPE = "2") OR (ADDTYPE = "3") then

		add2SQL = "INSERT INTO Admin ([TTID],[UN],[PW],[P1],[P2],[ACCLEVEL],[TITLE],[FN],[LN],[EMAIL],[DEPT]) Values ('" & RSCHECK("ID") & "','" & UN & "','" & PW & "','" & P1 & "','" & P2 & "','" & adminlevel & "','" & TITLE & "','" & FN & "','" & LN & "','" & EMAIL & "','" & DEPT & "')"
		else
		add2SQL = "INSERT INTO Users ([TTID],[UN],[PW],[P1],[P2],[TITLE],[FN],[LN],[EMAIL],[DEPT]) Values ('" & RSCHECK("ID") & "','" & UN & "','" & PW & "','" & P1 & "','" & P2 & "','" & TITLE & "','" & FN & "','" & LN & "','" & EMAIL & "','" & DEPT & "')"
		end if
		userconn.Execute(add2SQL)
		
		USER = RSCHECK("ID")
		
		RSCHECK.close
		set RSCHECK = nothing
	
		response.redirect "/pt/admin/staff.asp?id=6&good=1&user=" & USER
		else

		if (FN = "") OR (LN = "") OR (ENTITLEMENT = "") OR (UN = "") OR (PW = "") THEN
			response.redirect "/pt/admin/staff.asp?id=6&err=2"
		else
		end if

		RSCHECKSQL = "SELECT FN, LN FROM Timetables WHERE FN = '" & FN & "' AND LN = '" & LN & "'"
		Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
		RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
		
		if (ADDTYPE = "2") OR (ADDTYPE = "3") then
			RSCHECK2SQL = "SELECT UN FROM Admin WHERE UN = '" & UN & "'"
		else
			RSCHECK2SQL = "SELECT UN FROM Users WHERE UN = '" & UN & "'"
		end if
		
		Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
		RSCHECK2.Open RSCHECK2SQL, userconn, adopenkeyset, adlockoptimistic
		
		if RSCHECK.RECORDCOUNT => 1 then
			RSCHECK.close
			set RSCHECK = nothing
			response.redirect "/pt/admin/staff.asp?id=6&err=1"
		elseif RSCHECK2.RECORDCOUNT => 1 then
			RSCHECK.close
			set RSCHECK = nothing
			RSCHECK2.close
			set RSCHECK2 = nothing
			response.redirect "/pt/admin/staff.asp?id=6&err=1"
		else
			RSCHECK.close
			set RSCHECK = nothing
		end if
		
		addSQL = "INSERT INTO Timetables ([TITLE],[FN],[LN],[DEPT],[CATEGORY],[DEFROOM],[ENTITLEMENT]) Values ('" & TITLE & "', '" & FN & "', '" & LN & "', '" & DEPT & "', '" & CATEGORY & "', '" & DEFROOM & "', '" & ENTITLEMENT & "')"
		dataconn.Execute(addSQL)
	
		RSCHECKSQL = "SELECT ID FROM Timetables WHERE FN = '" & FN & "' AND LN = '" & LN & "' AND CATEGORY = '" & CATEGORY & "' AND DEFROOM = '" & DEFROOM & "' AND DEPT = " & DEPT
		Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
		RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic

			if ADDTYPE = "2" then
				adminlevel = 2
			elseif ADDTYPE = "3" then
				adminlevel = 1
			end if

		if (ADDTYPE = "2") OR (ADDTYPE = "3") then

		add2SQL = "INSERT INTO Admin ([TTID],[UN],[PW],[ACCLEVEL],[TITLE],[FN],[LN],[EMAIL],[DEPT]) Values ('" & RSCHECK("ID") & "','" & UN & "','" & PW & "','" & adminlevel & "','" & TITLE & "','" & FN & "','" & LN & "','" & EMAIL & "','" & DEPT & "')"
		else
		add2SQL = "INSERT INTO Users ([TTID],[UN],[PW],[TITLE],[FN],[LN],[EMAIL],[DEPT]) Values ('" & RSCHECK("ID") & "','" & UN & "','" & PW & "','" & TITLE & "','" & FN & "','" & LN & "','" & EMAIL & "','" & DEPT & "')"
		end if
		userconn.Execute(add2SQL)
		
		USER = RSCHECK("ID")
		
		RSCHECK.close
		set RSCHECK = nothing
	
		response.redirect "/pt/admin/staff.asp?id=6&good=1&user=" & USER
		end if
	else
		RSCHECKSQL = "SELECT FN, LN FROM Timetables WHERE FN = '" & FN & "' AND LN = '" & LN & "'"
		Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
		RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
		
		if RSCHECK.RECORDCOUNT => 1 then
			RSCHECK.close
			set RSCHECK = nothing
			response.redirect "/pt/admin/staff.asp?id=6&err=1"
		else
			RSCHECK.close
			set RSCHECK = nothing
		end if
	
		if (FN = "") OR (LN = "") OR (ENTITLEMENT = "") THEN
			response.redirect "/pt/admin/staff.asp?id=6&err=2"
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
	
		response.redirect "/pt/admin/staff.asp?id=6&good=1&user=" & USER
	end if

elseif addtype = "2" then

	if (request("type")) = "2" then
	
	daydow = request("dow")
	coverday = request("coverday")

	RSCHECKSQL = "SELECT ID, DEPT FROM Timetables"
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
	do until RSCHECK.EOF

		if (request("u_" & RSCHECK("ID"))) = "on" then
				RSCHECK2SQL = "SELECT DAY, DAYDATE, USER FROM Attendance WHERE DAYDATE = #" & SQLDATE(coverday) & "# AND DAY = " & daydow & " AND USER = " & RSCHECK("ID")
				Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
				RSCHECK2.Open RSCHECK2SQL, dataconn, adopenkeyset, adlockoptimistic
				
				if RSCHECK2.RECORDCOUNT > 0 then
				
				else
					addSQL = "INSERT INTO Attendance ([DAY],[DAYDATE],[USER],[DEPT]) Values ('" & daydow & "','" & coverday & "','" & RSCHECK("ID") & "','" & RSCHECK("DEPT") & "')"
					dataconn.Execute(addSQL)
				end if
				
				RSCHECK2.close
				set RSCHECK2 = nothing
				
		else
				RSCHECK2SQL = "SELECT DAY, USER FROM Attendance WHERE DAYDATE = #" & SQLDate(coverday) & "# AND DAY = " & daydow & " AND USER = " & RSCHECK("ID")
				Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
				RSCHECK2.Open RSCHECK2SQL, dataconn, adopenkeyset, adlockoptimistic
				
				if RSCHECK2.RECORDCOUNT > 0 then
					delSQL = "DELETE * FROM Attendance WHERE USER = " & RSCHECK("ID") & " AND DAY = " & daydow & " AND DAYDATE = #" & SQLDate(coverday) & "#"
					dataconn.Execute(delSQL)
					delSQL2 = "DELETE * FROM Cover WHERE FOR = " & RSCHECK("ID") & " AND DAY = " & daydow & " AND DAYDATE = #" & SQLDate(coverday) & "#"
					dataconn.Execute(delSQL2)
				else
				end if
				
				RSCHECK2.close
				set RSCHECK2 = nothing
		end if
	RSCHECK.MOVENEXT
	loop

	RSCHECK.close
	set RSCHECK = nothing

	response.redirect "/pt/admin/cover.asp?id=3&type=2&coverday=" & coverday & "&dow=" & daydow

	else
	
	DEPT = request("dept")

	RSCHECKSQL = "SELECT ID, DEPT FROM Timetables WHERE DEPT = " & request("dept")
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
	do until RSCHECK.EOF

		if (request("u_" & RSCHECK("ID"))) = "on" then
				RSCHECK2SQL = "SELECT DAY, USER FROM Attendance WHERE DAY = " & DOW & " AND USER = " & RSCHECK("ID")
				Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
				RSCHECK2.Open RSCHECK2SQL, dataconn, adopenkeyset, adlockoptimistic
				
				if RSCHECK2.RECORDCOUNT > 0 then
				
				else
					addSQL = "INSERT INTO Attendance ([DAY],[DAYDATE],[USER],[DEPT]) Values ('" & DOW & "','" & date() & "','" & RSCHECK("ID") & "','" & RSCHECK("DEPT") & "')"
					dataconn.Execute(addSQL)
				end if
				
				RSCHECK2.close
				set RSCHECK2 = nothing
				
		else
				RSCHECK2SQL = "SELECT DAY, USER FROM Attendance WHERE DAY = " & DOW & " AND USER = " & RSCHECK("ID")
				Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
				RSCHECK2.Open RSCHECK2SQL, dataconn, adopenkeyset, adlockoptimistic
				
				if RSCHECK2.RECORDCOUNT > 0 then
					delSQL = "DELETE * FROM Attendance WHERE USER = " & RSCHECK("ID")
					dataconn.Execute(delSQL)
					delSQL2 = "DELETE * FROM Cover WHERE FOR = " & RSCHECK("ID")
					dataconn.Execute(delSQL2)
				else
				end if
				
				RSCHECK2.close
				set RSCHECK2 = nothing
		end if
	RSCHECK.MOVENEXT
	loop

	RSCHECK.close
	set RSCHECK = nothing

	response.redirect "/pt/admin/cover.asp?id=3&type=1"

	end if

elseif addtype = "3" then

	DAYDATE = request("coverday")
	DAYDOW = request("dow")

	RSCHECKSQL = "SELECT * FROM Attendance WHERE DAY = " & DAYDOW & " AND DAYDATE = #" & SQLDate(DAYDATE) & "#"

	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic

	RSCHECK2SQL = "SELECT * FROM Periods WHERE ID = " & DAYDOW

	Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
	RSCHECK2.Open RSCHECK2SQL, dataconn, adopenkeyset, adlockoptimistic

	do until RSCHECK.EOF
		for i = 1 to RSCHECK2("TOTALS")
			covering = request(RSCHECK("USER")& "_" & i)
			if (covering = undefined) then
			elseif (left(covering,1) = "o") then
					coveringid = mid(covering,3)
					RSCHECK3SQL = "SELECT ID FROM COVER WHERE FOR = " & RSCHECK("USER") & " AND DAY = " & DAYDOW & " AND DAYDATE = #" & SQLDate(DAYDATE) & "# AND PERIOD = " & i

					Set RSCHECK3 = Server.CreateObject("Adodb.RecordSet")
					RSCHECK3.Open RSCHECK3SQL, dataconn, adopenkeyset, adlockoptimistic
					
					if RSCHECK3.RECORDCOUNT => 1 then
						delSQL = "DELETE * FROM Cover WHERE ID = " & RSCHECK3("ID")
						dataconn.Execute(delSQL)
					else
					end if
					
					RSCHECK3.close
					set RSCHECK3 = nothing

					RSCHECK4SQL = "SELECT ID FROM COVER WHERE COVERING = " & coveringid & " AND DAY = " & DAYDOW & " AND DAYDATE = #" & SQLDate(DAYDATE) & "# AND PERIOD = " & i

					Set RSCHECK4 = Server.CreateObject("Adodb.RecordSet")
					RSCHECK4.Open RSCHECK4SQL, dataconn, adopenkeyset, adlockoptimistic
					
					if RSCHECK4.RECORDCOUNT => 1 then
						errvar = 1
					else
						RSCHECK5SQL = "SELECT ID FROM Cover WHERE COVERING = " & coveringid

						Set RSCHECK5 = Server.CreateObject("Adodb.RecordSet")
						RSCHECK5.Open RSCHECK5SQL, dataconn, adopenkeyset, adlockoptimistic

						RSCHECK6SQL = "SELECT ENTITLEMENT FROM OCover WHERE ID = " & coveringid

						Set RSCHECK6 = Server.CreateObject("Adodb.RecordSet")
						RSCHECK6.Open RSCHECK6SQL, dataconn, adopenkeyset, adlockoptimistic
						
						entitlement = RSCHECK6("ENTITLEMENT") - RSCHECK5.RECORDCOUNT
						
						if entitlement > 0 then
							addSQL = "INSERT INTO Cover ([FOR],[COVERING],[DAY],[DAYDATE],[PERIOD],[OCOVER]) Values ('" & RSCHECK("USER") & "','" & coveringid & "','" & DAYDOW & "','" & DAYDATE & "','" & i & "','1')"
							dataconn.Execute(addSQL)
						else
							errvar = 2
						end if
					end if
			else
					coveringid = mid(covering,3)
					RSCHECK3SQL = "SELECT ID FROM COVER WHERE FOR = " & RSCHECK("USER") & " AND DAY = " & DAYDOW & " AND DAYDATE = #" & SQLDate(DAYDATE) & "# AND PERIOD = " & i

					Set RSCHECK3 = Server.CreateObject("Adodb.RecordSet")
					RSCHECK3.Open RSCHECK3SQL, dataconn, adopenkeyset, adlockoptimistic
					
					if RSCHECK3.RECORDCOUNT => 1 then
						delSQL = "DELETE * FROM Cover WHERE ID = " & RSCHECK3("ID")
						dataconn.Execute(delSQL)
					else
					end if
					
					RSCHECK3.close
					set RSCHECK3 = nothing

					RSCHECK4SQL = "SELECT ID FROM COVER WHERE COVERING = " & covering & " AND DAY = " & DAYDOW & " AND DAYDATE = #" & SQLDate(DAYDATE) & "# AND PERIOD = " & i

					Set RSCHECK4 = Server.CreateObject("Adodb.RecordSet")
					RSCHECK4.Open RSCHECK4SQL, dataconn, adopenkeyset, adlockoptimistic
					
					if RSCHECK4.RECORDCOUNT => 1 then
						errvar = 1
					else
						RSCHECK5SQL = "SELECT ID FROM Cover WHERE COVERING = " & covering

						Set RSCHECK5 = Server.CreateObject("Adodb.RecordSet")
						RSCHECK5.Open RSCHECK5SQL, dataconn, adopenkeyset, adlockoptimistic

						RSCHECK6SQL = "SELECT ENTITLEMENT FROM Timetables WHERE ID = " & covering

						Set RSCHECK6 = Server.CreateObject("Adodb.RecordSet")
						RSCHECK6.Open RSCHECK6SQL, dataconn, adopenkeyset, adlockoptimistic
						
						entitlement = RSCHECK6("ENTITLEMENT") - RSCHECK5.RECORDCOUNT
						
						if entitlement > 0 then
							addSQL = "INSERT INTO Cover ([FOR],[COVERING],[DAY],[DAYDATE],[PERIOD],[OCOVER]) Values ('" & RSCHECK("USER") & "','" & covering & "','" & DAYDOW & "','" & DAYDATE & "','" & i & "','0')"
							dataconn.Execute(addSQL)
						else
							errvar = 2
						end if
					end if
			end if
		next	
	RSCHECK.MOVENEXT
	loop

	if (request("type")) = "2" then
		if (errvar = "") then
			response.redirect "/pt/admin/cover.asp?id=4&type=2&coverday=" & DAYDATE & "&dow=" & DAYDOW
		else
			response.redirect "/pt/admin/cover.asp?id=4&type=2&coverday=" & DAYDATE & "&dow=" & DAYDOW & "&err=" & errvar
		end if
	else
		if (errvar = "") then
			response.redirect "/pt/admin/cover.asp?id=4&type=1&coverday=" & DAYDATE & "&dow=" & DAYDOW
		else
			response.redirect "/pt/admin/cover.asp?id=4&type=1&coverday=" & DAYDATE & "&dow=" & DAYDOW & "&err=" & errvar
		end if
	end if

	RSCHECK.close
	set RSCHECK = nothing

elseif addtype = "4" then

	ROOM = request("ROOM")

	RSCHECKSQL = "SELECT * FROM Rooms WHERE ROOMNO = '" & ROOM & "'"
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic

	if RSCHECK.RECORDCOUNT => 1 then

		RSCHECK.close
		set RSCHECK = nothing
		response.redirect "/pt/admin/rooms.asp?id=2&err=2"
	else

		RSCHECK.close
		set RSCHECK = nothing
		if ROOM = "" then

			response.redirect "/pt/admin/rooms.asp?id=2&err=1"

		else

			addSQL = "INSERT INTO Rooms ([ROOMNO]) Values ('" & ROOM & "')"
			dataconn.Execute(addSQL)

		end if

		response.redirect "/pt/admin/rooms.asp?id=2&gd=1"

	end if

elseif addtype = "5" then

	UID = request("UID")
	UN = request("UN")
	PW = request("PW")
	PWV = request("PWV")
	P1 = request("P1")
	P2 = request("P2")
	EMAIL = request("EMAIL") & "@" & var_emaildomain1
	
	if var_est_enabled_pin = "1" then
	
		if (UN = "") OR (PW = "") OR (P1 = "") OR (P2 = "") OR (PWV = "") OR (EMAIL = "") then
			response.redirect "/pt/signup.asp?id=4&uid=" & UID & "&dept=" & request("dept") & "&err=1"
		else
			if (PW <> PWV) then
				response.redirect "/pt/signup.asp?id=4&uid=" & UID & "&dept=" & request("dept") & "&err=3"
			else
				RSCHECKSQL = "SELECT * FROM Users WHERE UN = '" & UN & "'"
				Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
				RSCHECK.Open RSCHECKSQL, userconn, adopenkeyset, adlockoptimistic
				
				if RSCHECK.RECORDCOUNT => 1 then
					RSCHECK.close
					set RSCHECK = nothing
					response.redirect "/pt/signup.asp?id=4&uid=" & UID & "&dept=" & request("dept") & "&err=2"
				else
					RSCHECK.close
					set RSCHECK = nothing

					RSCHECK2SQL = "SELECT * FROM Timetables WHERE ID = " & UID
					Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
					RSCHECK2.Open RSCHECK2SQL, dataconn, adopenkeyset, adlockoptimistic
					
					if RSCHECK2.RECORDCOUNT = 0 then
						response.redirect "/pt/signup.asp?id=4&uid=" & UID & "&dept=" & request("dept") & "&err=4"

						RSCHECK2.close
						set RSCHECK2 = nothing
					else
						addSQL = "INSERT INTO Users ([TTID],[UN],[PW],[P1],[P2],[TITLE],[FN],[LN],[EMAIL],[DEPT]) Values ('" & UID & "','" & UN & "','" & PW & "','" & P1 & "','" & P2 & "','" & RSCHECK2("TITLE") & "','" & RSCHECK2("FN") & "','" & RSCHECK2("LN") & "','" & EMAIL & "','" & RSCHECK2("DEPT") & "')"
						userconn.Execute(addSQL)
						
						RSCHECK2.close
						set RSCHECK2 = nothing

						response.redirect "/pt/signup.asp?id=5"
					end if
				end if
			end if
		end if

	else
	
		if (UN = "") OR (PW = "") OR (PWV = "") OR (EMAIL = "") then
			response.redirect "/pt/signup.asp?id=4&uid=" & UID & "&dept=" & request("dept") & "&err=1"
		else
			if (PW <> PWV) then
				response.redirect "/pt/signup.asp?id=4&uid=" & UID & "&dept=" & request("dept") & "&err=3"
			else
				RSCHECKSQL = "SELECT * FROM Users WHERE UN = '" & UN & "'"
				Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
				RSCHECK.Open RSCHECKSQL, userconn, adopenkeyset, adlockoptimistic
				
				if RSCHECK.RECORDCOUNT => 1 then
					RSCHECK.close
					set RSCHECK = nothing
					response.redirect "/pt/signup.asp?id=4&uid=" & UID & "&dept=" & request("dept") & "err=2"
				else
					RSCHECK.close
					set RSCHECK = nothing

					RSCHECK2SQL = "SELECT * FROM Timetables WHERE ID = " & UID
					Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
					RSCHECK2.Open RSCHECK2SQL, dataconn, adopenkeyset, adlockoptimistic
					
					if RSCHECK2.RECORDCOUNT = 0 then
						response.redirect "/pt/signup.asp?id=4&uid=" & UID & "&dept=" & request("dept") & "&err=4"

						RSCHECK2.close
						set RSCHECK2 = nothing
					else
						addSQL = "INSERT INTO Users ([TTID],[UN],[PW],[TITLE],[FN],[LN],[EMAIL],[DEPT]) Values ('" & UID & "','" & UN & "','" & PW & "','" & RSCHECK2("TITLE") & "','" & RSCHECK2("FN") & "','" & RSCHECK2("LN") & "','" & EMAIL & "','" & RSCHECK2("DEPT") & "')"
						userconn.Execute(addSQL)
						
						RSCHECK2.close
						set RSCHECK2 = nothing
						
						response.redirect "/pt/signup.asp?id=5"
					end if
				end if
			end if
		end if

	end if

elseif addtype = "6" then

	UID = request("UID")
	UN = request("UN")
	PW = request("PW")
	P1 = request("P1")
	P2 = request("P2")
	EMAIL = request("EMAIL") & "@" & var_emaildomain1
	CREATE = request("TYPE")
	
	if var_est_enabled_pin = "1" then
		if (UN = "") OR (PW = "") OR (P1 = "") OR (P2 = "") OR (EMAIL = "") then
			response.redirect "/pt/admin/staff.asp?id=8&uid=" & uid & "&err=1"
		else
			if (CREATE = "2") OR (CREATE = "3") then
				RSCHECKSQL = "SELECT * FROM Admin WHERE UN = '" & UN & "'"
			else
				RSCHECKSQL = "SELECT * FROM Users WHERE UN = '" & UN & "'"
			end if

			Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
			RSCHECK.Open RSCHECKSQL, userconn, adopenkeyset, adlockoptimistic
			
			if RSCHECK.RECORDCOUNT => 1 then
				RSCHECK.close
				set RSCHECK = nothing
					response.redirect "/pt/admin/staff.asp?id=8&uid=" & UID & "&err=2"
			else

				RSCHECK2SQL = "SELECT * FROM Timetables WHERE ID = " & UID
				Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
				RSCHECK2.Open RSCHECK2SQL, dataconn, adopenkeyset, adlockoptimistic
					
				if RSCHECK2.RECORDCOUNT = 0 then
					RSCHECK2.close
					set RSCHECK2 = nothing
					response.redirect "/pt/admin/staff.asp.asp?id=8&uid=" & UID & "&err=3"
				else
					if CREATE = "1" then
						addSQL = "INSERT INTO Users ([TTID],[UN],[PW],[P1],[P2],[TITLE],[FN],[LN],[EMAIL],[DEPT]) Values ('" & UID & "','" & UN & "','" & PW & "','" & P1 & "','" & P2 & "','" & RSCHECK2("TITLE") & "','" & RSCHECK2("FN") & "','" & RSCHECK2("LN") & "','" & EMAIL & "','" & RSCHECK2("DEPT") & "')"
					elseif CREATE = "2" then
						addSQL = "INSERT INTO Admin ([TTID],[UN],[PW],[P1],[P2],[ACCLEVEL],[TITLE],[FN],[LN],[EMAIL],[DEPT]) Values ('" & UID & "','" & UN & "','" & PW & "','" & P1 & "','" & P2 & "','2','" & RSCHECK2("TITLE") & "','" & RSCHECK2("FN") & "','" & RSCHECK2("LN") & "','" & EMAIL & "','" & RSCHECK2("DEPT") & "')"
					elseif CREATE = "3" then
						addSQL = "INSERT INTO Admin ([TTID],[UN],[PW],[P1],[P2],[ACCLEVEL],[TITLE],[FN],[LN],[EMAIL],[DEPT]) Values ('" & UID & "','" & UN & "','" & PW & "','" & P1 & "','" & P2 & "','1','" & RSCHECK2("TITLE") & "','" & RSCHECK2("FN") & "','" & RSCHECK2("LN") & "','" & EMAIL & "','" & RSCHECK2("DEPT") & "')"
					end if
					userconn.Execute(addSQL)

					RSCHECK2.close
					set RSCHECK2 = nothing
					
					response.redirect "/pt/admin/staff.asp?id=8&gd=1"
				end if			
			end if
		end if
	else
		if (UN = "") OR (PW = "") OR (EMAIL = "") then
			response.redirect "/pt/admin/staff.asp?id=8&uid=" & uid & "&err=1"
		else
			if (CREATE = "2") OR (CREATE = "3") then
				RSCHECKSQL = "SELECT * FROM Admin WHERE UN = '" & UN & "'"
			else
				RSCHECKSQL = "SELECT * FROM Users WHERE UN = '" & UN & "'"
			end if

			Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
			RSCHECK.Open RSCHECKSQL, userconn, adopenkeyset, adlockoptimistic
			
			if RSCHECK.RECORDCOUNT => 1 then
				RSCHECK.close
				set RSCHECK = nothing
					response.redirect "/pt/admin/staff.asp?id=8&uid=" & UID & "&err=2"
			else

				RSCHECK2SQL = "SELECT * FROM Timetables WHERE ID = " & UID
				Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
				RSCHECK2.Open RSCHECK2SQL, dataconn, adopenkeyset, adlockoptimistic
					
				if RSCHECK2.RECORDCOUNT = 0 then
					RSCHECK2.close
					set RSCHECK2 = nothing
					response.redirect "/pt/admin/staff.asp.asp?id=8&uid=" & UID & "&err=3"
				else
					if CREATE = "1" then
						addSQL = "INSERT INTO Users ([TTID],[UN],[PW],[TITLE],[FN],[LN],[EMAIL],[DEPT]) Values ('" & UID & "','" & UN & "','" & PW & "','" & RSCHECK2("TITLE") & "','" & RSCHECK2("FN") & "','" & RSCHECK2("LN") & "','" & EMAIL & "','" & RSCHECK2("DEPT") & "')"
					elseif CREATE = "2" then
						addSQL = "INSERT INTO Admin ([TTID],[UN],[PW],[ACCLEVEL],[TITLE],[FN],[LN],[EMAIL],[DEPT]) Values ('" & UID & "','" & UN & "','" & PW & "','2','" & RSCHECK2("TITLE") & "','" & RSCHECK2("FN") & "','" & RSCHECK2("LN") & "','" & EMAIL & "','" & RSCHECK2("DEPT") & "')"
					elseif CREATE = "3" then
						addSQL = "INSERT INTO Admin ([TTID],[UN],[PW],[ACCLEVEL],[TITLE],[FN],[LN],[EMAIL],[DEPT]) Values ('" & UID & "','" & UN & "','" & PW & "','1','" & RSCHECK2("TITLE") & "','" & RSCHECK2("FN") & "','" & RSCHECK2("LN") & "','" & EMAIL & "','" & RSCHECK2("DEPT") & "')"
					end if
					userconn.Execute(addSQL)

					RSCHECK2.close
					set RSCHECK2 = nothing
					
					response.redirect "/pt/admin/staff.asp?id=8&gd=1"
				end if			
			end if
		end if
	end if

elseif addtype = "7" then

'DEPT ADD

elseif addtype = "8" then

	TITLE = request("TITLE")
	FN = replace(request("FN"), "'", "''")
	LN = replace(request("LN"), "'", "''")
	ENTITLEMENT = request("ENTITLEMENT")

	if (FN = "") or (LN = "") or (ENTITLEMENT = "") or (isnumeric(ENTITLEMENT) = false) then
		response.redirect "/pt/admin/ocover.asp?id=1&err=1"
	else
		RSCHECKSQL = "SELECT * FROM OCover WHERE TITLE = '" & TITLE & "' AND FN = '" & FN & "' AND LN = '" & LN & "'"
		Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
		RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
		if RSCHECK.RECORDCOUNT => "1" then
			response.redirect "/pt/admin/ocover.asp?id=1&err=2"
		else		
			addSQL = "INSERT INTO OCover ([LN],[FN],[TITLE],[ENTITLEMENT]) Values ('" & LN & "','" & FN & "','" & TITLE & "','" & ENTITLEMENT & "')"
			dataconn.Execute(addSQL)
			response.redirect "/pt/admin/ocover.asp?id=1"
		end if
		RSCHECK.close
		set RSCHECK = nothing
	end if
else

end if
%>