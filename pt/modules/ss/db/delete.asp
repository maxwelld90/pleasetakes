<!--
Server-ML.co.uk PleaseTake System
Database Management Subsystem - Delete
Copyright (c) Server-ML.co.uk 2006

deltype(1) = Deleting A Member Of Staff
deltype(2) = Deleting A Cover (From Wizard, Deselecting)
deltype(3) = Deleting A Room
deltype(4) = Deleting An Access Account
deltype(5) = Deleting A Department
deltype(6) = Deleting An Outside Cover Member
deltype(7) = Clearing Cover Requests For A Day (Wizard)
deltype(8) = Deleting A Cover Request (Cover Summary/Popup)
-->

<!--#include virtual="/pt/modules/ss/p_s.inc"-->
<%
deltype = request("deltype")

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

if deltype = "1" then
	USER = request("USER")
	
	RSCHECKSQL = "SELECT USER FROM Attendance WHERE USER = " & USER
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
	
	delSQL = "DELETE * FROM Attendance WHERE USER = " & USER
	dataconn.Execute(delSQL)			
	
	delSQL = "DELETE * FROM Timetables WHERE ID = " & USER
	dataconn.Execute(delSQL)
	
	RSCHECK.close
	set RSCHECK = nothing
	
	response.redirect "/pt/admin/staff.asp?id=7&good=1"
	
elseif deltype="2" then
	delSQL = "DELETE * FROM Cover WHERE ID = " & request("ID")
	dataconn.Execute(delSQL)
	
	if (request("type")) = "2" then
		response.redirect "/pt/admin/cover.asp?id=4&type=2&coverday=" & request("coverday") & "&dow=" & request("dow")
	else
		response.redirect "/pt/admin/cover.asp?id=4&type=1&coverday=" & request("coverday") & "&dow=" & request("dow")
	end if

elseif deltype = "3" then

	ROOM = request("ROOM")

	RSCHECKSQL = "SELECT * FROM Rooms WHERE ID = " & ROOM
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
	
	if RSCHECK.RECORDCOUNT = 0 then
		response.redirect "/pt/admin/rooms.asp?id=4&err=1"
	else
		
		delSQL = "DELETE * FROM Rooms WHERE ID = " & ROOM
		dataconn.Execute(delSQL)
	
		response.redirect "/pt/admin/rooms.asp?id=4&gd=1"
	end if

elseif deltype = "4" then

	UID = request("UID")

	RSCHECKSQL = "SELECT * FROM Users WHERE TTID = " & UID
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, userconn, adopenkeyset, adlockoptimistic
	
	if RSCHECK.RECORDCOUNT = 0 then
		RSCHECK2SQL = "SELECT * FROM Admin WHERE TTID = " & UID
		Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
		RSCHECK2.Open RSCHECK2SQL, userconn, adopenkeyset, adlockoptimistic
		
		if RSCHECK2.RECORDCOUNT = 0 then
			RSCHECK.close
			set RSCHECK = nothing
			RSCHECK2.close
			set RSCHECK2 = nothing
				response.redirect "/pt/admin/staff.asp?id=2"
		else
			delSQL = "DELETE * FROM Admin WHERE TTID = " & UID
			userconn.Execute(delSQL)

			RSCHECK.close
			set RSCHECK = nothing
			RSCHECK2.close
			set RSCHECK2 = nothing

			response.redirect "/pt/admin/staff.asp?id=8&manage=1&del=2"
		end if
	else
		delSQL = "DELETE * FROM Users WHERE TTID = " & UID
		dataconn.Execute(delSQL)

		RSCHECK.close
		set RSCHECK = nothing
		
		response.redirect "/pt/admin/staff.asp?id=8&manage=1&del=2"
	end if

elseif deltype = "5" then

	'DEL DEPT

elseif deltype = "6" then

		uid = request("UID")

		if (isnumeric(uid) = false) then
			response.redirect "/pt/admin/ocover.asp?id=1"
		else
			if uid = "" then
				response.redirect "/pt/admin/ocover.asp?id=1"
			else
				delSQL = "DELETE * FROM OCover WHERE ID = " & UID
				dataconn.Execute(delSQL)
				response.redirect "/pt/admin/ocover.asp?id=1"
			end if
		end if
elseif deltype = "7" then
	delSQL = "DELETE * FROM Cover WHERE DAYDATE = #" & SQLDate(request("coverday")) & "# AND DAY = " & request("dow")
	dataconn.execute(delSQL)
	
	response.redirect "/pt/admin/cover.asp?id=4&type=" & request("type") & "&coverday=" & request("coverday") & "&dow=" & request("dow")

elseif deltype = "8" then

	delSQL = "DELETE * FROM Cover WHERE ID = " & request("cover")
	dataconn.execute(delSQL)
%>
	<script type="text/javascript">opener.location.reload(true); self.close();</script>
<%
else

end if
%>