<!--
Server-ML.co.uk PleaseTake System
User Subsystem - Logout
Copyright (c) Server-ML.co.uk 2006

logintype(1) = Standard User
logintype(2) = Administrator
-->

<%
logouttype = request("id")

if logouttype = "1" then

	session("sess_loggedins") = False
	session("sess_acclevel") = ""
	session("sess_un") = ""
	session("sess_fn") = ""
	session("sess_ln") = ""
	session("sess_email") = ""
	session("sess_dept") = ""
	session.abandon
	response.redirect "/pt/default.asp?id=7"

elseif logouttype = "2" then

	session("sess_loggedina") = False
	session("sess_acclevel") = ""
	session("sess_un") = ""
	session("sess_fn") = ""
	session("sess_ln") = ""
	session("sess_email") = ""
	session("sess_dept") = ""
	session.abandon
	response.redirect "/pt/default.asp?id=8"

end if
%>