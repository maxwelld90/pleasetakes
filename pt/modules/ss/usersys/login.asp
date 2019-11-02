<!--
Server-ML.co.uk PleaseTake System
User Subsystem - Login
Copyright (c) Server-ML.co.uk 2006

logintype(1) = Standard User
logintype(2) = Administrator
-->

<!--#include virtual="/pt/modules/ss/p_s.inc"-->
<%
logintype = request("id")
un = request("un")
pw = request("pw")
p1 = request("p1")
p2 = request("p2")

if logintype = "1" then

	if var_est_enabled <> 1 then
		response.redirect "/pt/default.asp?id=1"
	else

	if var_est_enabled_pin = 1 then
	
	Set conn = server.createobject("adodb.connection")
	DBopen = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='D:\Websites\Data\PleaseTakes\users.mdb'"
	conn.Open DBopen

	if (p1 = "") AND (p2 = "") then

	RSUSERSQL = "SELECT * FROM Users WHERE UN = '" & un & "' AND PW = '" & pw & "'"

	else
	
	RSUSERSQL = "SELECT * FROM Users WHERE UN = '" & un & "' AND PW = '" & pw & "' AND P1 = '" & p1 & "' AND P2 = '" & p2 & "'"

	end if
	
	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, conn, adopenkeyset, adlockoptimistic
	
	if RSUSER.recordcount = 0 then
	RSUSER.close
	set RSUSER = nothing
	response.redirect "/pt/default.asp?id=3"
	else
	session("sess_loggedins") = True
	session("sess_acclevel") = 1
	session("sess_un") = RSUSER("UN")
	session("sess_fn") = RSUSER("FN")
	session("sess_ln") = RSUSER("LN")
	session("sess_email") = RSUSER("EMAIL")
	session("sess_dept") = RSUSER("DEPT")
	session("sess_lastin") = RSUSER("LASTLOGIN")
	session("sess_ttid") = RSUSER("TTID")
	
	loginup = "UPDATE users SET LASTLOGIN = '" & dbxml_date & ", " & dbxml_time & "' WHERE UN = '" & RSUSER("un") & "'"
	conn.execute (loginup)
	
	RSUSER.close
	set RSUSER = nothing

	response.redirect "/pt/std/default.asp"

	end if
	
	else
	
	Set conn = server.createobject("adodb.connection")
	DBopen = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='D:\Websites\Data\PleaseTakes\users.mdb'"
	conn.Open DBopen
	
	RSUSERSQL = "SELECT * FROM Users WHERE UN = '" & un & "' AND PW = '" & pw & "'"
	
	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, conn, adopenkeyset, adlockoptimistic
	
	if RSUSER.recordcount = 0 then
	RSUSER.close
	set RSUSER = nothing
	response.redirect "/pt/default.asp?id=3"
	else
	session("sess_loggedins") = True
	session("sess_acclevel") = 1
	session("sess_un") = RSUSER("UN")
	session("sess_fn") = RSUSER("FN")
	session("sess_ln") = RSUSER("LN")
	session("sess_email") = RSUSER("EMAIL")
	session("sess_dept") = RSUSER("DEPT")
	session("sess_lastin") = RSUSER("LASTLOGIN")
	session("sess_ttid") = RSUSER("TTID")

	loginup = "UPDATE users SET LASTLOGIN = '" & dbxml_date & ", " & dbxml_time & "' WHERE UN = '" & RSUSER("un") & "'"
	conn.execute (loginup)

	RSUSER.close
	set RSUSER = nothing

	if (var_firstlogin = "1") AND (un <> settingsXML.documentElement.childNodes.item(0).childNodes.item(4).getAttribute("firstloginacc")) then

	response.redirect "/pt/default.asp?id=9"

	else

	response.redirect "/pt/std/default.asp?id=1"

	end if

	end if
	end if

	end if

elseif logintype = "2" then

	if var_est_enabled_pin = 1 then
	
	Set conn = server.createobject("adodb.connection")
	DBopen = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='D:\Websites\Data\PleaseTakes\users.mdb'"
	conn.Open DBopen
	
	if (p1 = "") AND (p2 = "") then

	RSUSERSQL = "SELECT * FROM Admin WHERE UN = '" & un & "' AND PW = '" & pw & "'"

	else
	
	RSUSERSQL = "SELECT * FROM Admin WHERE UN = '" & un & "' AND PW = '" & pw & "' AND P1 = '" & p1 & "' AND P2 = '" & p2 & "'"

	end if
	
	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, conn, adopenkeyset, adlockoptimistic
	
	if RSUSER.recordcount = 0 then
	RSUSER.close
	set RSUSER = nothing
	response.redirect "/pt/default.asp?id=4"
	else

	if (var_est_enabled <> 1) AND (RSUSER("ACCLEVEL") <> 1) then
		response.redirect "/pt/default.asp?id=2"
	else

	session("sess_loggedina") = True
	session("sess_acclevel") = 2
	session("sess_adminlevel") = RSUSER("ACCLEVEL")
	session("sess_un") = RSUSER("UN")
	session("sess_fn") = RSUSER("FN")
	session("sess_ln") = RSUSER("LN")
	session("sess_email") = RSUSER("EMAIL")
	session("sess_dept") = RSUSER("DEPT")
	session("sess_lastin") = RSUSER("LASTLOGIN")
	session("sess_ttid") = RSUSER("TTID")

	loginup = "UPDATE admin SET LASTLOGIN = '" & dbxml_date & ", " & dbxml_time & "' WHERE UN = '" & RSUSER("un") & "'"
	conn.execute (loginup)

	RSUSER.close
	set RSUSER = nothing

	if var_firstlogin = "1" then

	response.redirect "/pt/admin/setup.asp?id=1"

	else

	response.redirect "/pt/admin/default.asp?id=1"

	end if

	end if

	end if
	
	else
	
	Set conn = server.createobject("adodb.connection")
	DBopen = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='D:\Websites\Data\PleaseTakes\users.mdb'"
	conn.Open DBopen
	
	RSUSERSQL = "SELECT * FROM Admin WHERE UN = '" & un & "' AND PW = '" & pw & "'"
	
	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, conn, adopenkeyset, adlockoptimistic
	
	if RSUSER.recordcount = 0 then
	RSUSER.close
	set RSUSER = nothing
	response.redirect "/pt/default.asp?id=4"
	else

	if (var_est_enabled <> 1) AND (RSUSER("ACCLEVEL") <> 1) then
		response.redirect "/pt/default.asp?id=2"
	else

	session("sess_loggedina") = True
	session("sess_acclevel") = 2
	session("sess_adminlevel") = RSUSER("ACCLEVEL")
	session("sess_un") = RSUSER("UN")
	session("sess_fn") = RSUSER("FN")
	session("sess_ln") = RSUSER("LN")
	session("sess_email") = RSUSER("EMAIL")
	session("sess_dept") = RSUSER("DEPT")
	session("sess_lastin") = RSUSER("LASTLOGIN")
	session("sess_ttid") = RSUSER("TTID")

	loginup = "UPDATE admin SET LASTLOGIN = '" & dbxml_date & ", " & dbxml_time & "' WHERE UN = '" & RSUSER("un") & "'"
	conn.execute (loginup)

	RSUSER.close
	set RSUSER = nothing


	if (var_firstlogin = "1") AND (un = settingsXML.documentElement.childNodes.item(0).childNodes.item(4).getAttribute("firstloginacc")) then

	response.redirect "/pt/admin/setup.asp?id=1"

	elseif (var_firstlogin = "1") AND (un <> settingsXML.documentElement.childNodes.item(0).childNodes.item(4).getAttribute("firstloginacc")) then

	response.redirect "/pt/default.asp?id=10"

	else

	response.redirect "/pt/admin/default.asp?id=1"

	end if
	end if
	end if
	end if

end if
%>