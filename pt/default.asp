<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" >

<!--#include virtual="/pt/modules/ss/p_s.inc"-->
<%
logintype = request("id")
var_comment = settingsXML.documentElement.childNodes.item(1).childNodes.item(0).text
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="ROBOTS" content="NOINDEX">
<link rel="stylesheet" type="text/css" href="modules/css/login.css">
<%if var_est_enabled <> 1 then%><script language="javascript" type="text/javascript" src="/pt/modules/js/login.js"></script><%else%><script language="javascript" type="text/javascript" src="/pt/modules/js/login.js"></script><%end if%>
<title><%=var_ptitle%></title>
</head>

<body<%if var_est_enabled <> 1 then%><%if (logintype = "2") OR (logintype = "4") OR (logintype = "6") OR (logintype = "8") OR (logintype = "10") then%> onload="load();"<%else%> onload="offlineload();"<%end if%><%else%> onload="load();"<%end if%>>

<form name="frm" method="post" action="<%if var_est_enabled <> 1 then%><%if (logintype = "2") OR (logintype = "4") OR (logintype = "6") OR (logintype = "8") OR (logintype = "10") then%>/pt/modules/ss/usersys/login.asp?id=2<%else%>#<%end if%><%else%>/pt/modules/ss/usersys/login.asp?id=<%if (logintype = "2") OR (logintype = "4") OR (logintype = "6") OR (logintype = "8") OR (logintype = "10") then%>2<%else%>1<%end if%><%end if%>">
<table class="main" cellpadding="0" cellspacing="0">
	<tr>
		<td class="tl">
			<img src="/pt/media/login/logo.png" border="0" alt="<%=var_est_full%>">
			<br><span class="title"><%=var_pname%></span>
			<br><span class="ver">Version <%=var_ver%> (<%=var_est_short%>)</span>
			<br><span class="com"><%=display_date%>, <%=display_time%></span>
			<br><span class="com"><%=var_comment%></span><br>
		</td>
		<td class="p" rowspan="2"></td>
		<td class="c" rowspan="2">
			<div class="c_t"></div>
			<div class="c_b"></div>
		</td>
		<td class="p" rowspan="2"></td>
		<td class="tr">
			<div class="frmcont">
			<table class="formcontainer" cellpadding="0" cellspacing="0">
				<tr>
					<td class="ltitle"><%if var_est_enabled <> 1 then%><%if (logintype = "2") OR (logintype = "4") OR (logintype = "6") OR (logintype = "8") OR (logintype = "10") then%><%=var_usernames_admin%> Login<%else%>System Offline<%end if%><%else%><%if (logintype = "2") OR (logintype = "4") OR (logintype = "6") OR (logintype = "8") OR (logintype = "10") then%><%=var_usernames_admin%><%else%><%=var_usernames_std%><%end if%> Login<%end if%></td>
				</tr>
				<tr>
					<td>
					<table class="formcontainer" cellpadding="0" cellspacing="0">
						<tr>
							<td class="linfo" style="width: 68px;">Username</td>
							<td class="linfo"><input class="b_std" type="text" name="un" size="20" maxlength="20"<%if var_est_enabled <> 1 then%><%if (logintype = "2") OR (logintype = "4") OR (logintype = "6") OR (logintype = "8") OR (logintype = "10") then%><%else%> disabled<%end if%><%else%><%if (logintype = "2") OR (logintype = "4") OR (logintype = "6") OR (logintype = "8") OR (logintype = "10") then%><%else%><%end if%><%end if%>></td>
						</tr>
						<tr>
							<td class="linfo" style="width: 68px;">Password</td>
							<td class="linfo"><input class="b_pwd" type="password" name="pw" size="20" maxlength="20"<%if var_est_enabled <> 1 then%><%if (logintype = "2") OR (logintype = "4") OR (logintype = "6") OR (logintype = "8") OR (logintype = "10") then%><%else%> disabled<%end if%><%else%><%if (logintype = "2") OR (logintype = "4") OR (logintype = "6") OR (logintype = "8") OR (logintype = "10") then%><%else%><%end if%><%end if%>></td>
						</tr>
						<%
						if var_est_enabled_pin = 1 then
						%>
						<tr>
							<td class="linfo" style="width: 68px;">PIN</td>
							<td class="linfo"><input class="b_pwd" type="password" name="p1" size="3" maxlength="3" onkeyup="pjump();"<%if var_est_enabled <> 1 then%><%if (logintype = "2") OR (logintype = "4") OR (logintype = "6") OR (logintype = "8") OR (logintype = "10") then%><%else%> disabled<%end if%><%else%><%if (logintype = "2") OR (logintype = "4") OR (logintype = "6") OR (logintype = "8") OR (logintype = "10") then%><%else%><%end if%><%end if%>>&nbsp;/&nbsp;<input class="b_pwd" type="password" name="p2" size="3" maxlength="3"<%if var_est_enabled <> 1 then%><%if (logintype = "2") OR (logintype = "4") OR (logintype = "6") OR (logintype = "8") OR (logintype = "10") then%><%else%> disabled<%end if%><%else%><%if (logintype = "2") OR (logintype = "4") OR (logintype = "6") OR (logintype = "8") OR (logintype = "10") then%><%else%><%end if%><%end if%>></td>
						</tr>
						<%
						else
						end if
						%>
					</table>
					</td>
				</tr>
				<tr>
					<td class="brow">
						<%if var_est_enabled <> 1 then%><%if (logintype = "2") OR (logintype = "4") OR (logintype = "6") OR (logintype = "8") OR (logintype = "10") then%><a title="Click Here To Login Using The Details Above."><input class="b" type="submit" value="Login" name="go"></a>&nbsp;&nbsp;<a title="If You Made A Mistake, Click This Button To Clear The Fields Above."><input class="b" type="reset" value="Clear" name="rs" onmouseup="load();"></a><%else%><a title="The System Is Offline, So You Cannot Login."><input class="b" type="submit" value="Login" name="go" disabled></a>&nbsp;&nbsp;<a title="Cannot Clear The Above Fields."><input class="b" type="reset" value="Clear" name="rs" onmouseup="load();" disabled></a><%end if%><%else%><a title="Click Here To Login Using The Details Above."><input class="b" type="submit" value="Login" name="go"></a>&nbsp;&nbsp;<a title="If You Made A Mistake, Click This Button To Clear The Fields Above."><input class="b" type="reset" value="Clear" name="rs" onmouseup="load();"></a><%end if%>
					</td>
				</tr>
				<tr>
					<td class="ltext">
<%if var_est_enabled <> 1 then
	if (logintype = "2") OR (logintype = "4") OR (logintype = "6") OR (logintype = "8") OR (logintype = "10") then
%>
	Please Enter Your Details Above And Hit Login.<br>
	Only Full-Access Administrators May Login.
<%
	else
%>
	Sorry, The System Administrator Has Disabled The System.<br>
	Please Try Coming Back Later.
<%
	end if
else
	if (logintype = "3") OR (logintype = "4") then%>
	Sorry, One Or More Of The Details You Entered Are Invalid!<br>
	Please Try Again.
	<%elseif (logintype = "5") OR (logintype = "6") then%>
	Sorry, You Must Login First Before You Can Use This System.<br>
	Please Enter Your Details Above, And Hit Login.
	<%elseif (logintype = "7") OR (logintype = "8") then%>
	You Have Been Successfully Logged Out!<br>
	Thanks For Using The System!
	<%elseif (logintype = "9") OR (logintype = "10") then%>
	Sorry, You May Not Login Until The System Has Been Set Up.<br>
	Please Try Again Later To See If It Has.
	<%else%>
	<script language="javascript" type="text/javascript">document.write(daymsg)</script>, And Welcome To <%=var_pname%>!<br>Please Enter Your Details And Hit Login.
	<%end if
end if
%>					</td>
				</tr>
			</table>
			</div>		
		</td>
	</tr>
	<tr>
		<td class="bl"><a href="http://validator.w3.org/check?uri=referer"><img src="/pt/media/login/w3c.gif" border="0" alt="Using W3C Recommended HTML 4.01 Transitional And CSS2"></a><a href="http://www.server-ml.co.uk/"><img src="/pt/media/login/sml.gif" border="0" alt="Copyright Server-ML.co.uk 2006"></a></td>
		<td class="br">
<%
if var_est_enabled <> 1 then
%>
<a href="default.asp?id=2" title="Click Here To Switch To The <%=var_usernames_admin%> Login Screen.">Admin Login</a>
<%
else

Set menuXML = Server.CreateObject("Microsoft.XMLDOM")
Set menuXML = settingsXML.documentElement.childNodes.item(1).childNodes.item(1).childNodes

if (logintype = "2") OR (logintype = "4") OR (logintype = "6") OR (logintype = "8") OR (logintype = "10") then
	if menuXML.length => 1 then%>
			<a href="default.asp?id=1" title="Click Here To Switch To The <%=var_usernames_std%> Login Screen."><%=var_usernames_std%> Login</a> :: 
	<%
	else
	%>
			<a href="default.asp?id=1" title="Click Here To Switch To The <%=var_usernames_std%> Login Screen."><%=var_usernames_std%> Login</a>
	<%
	end if
	%>
<%else
	if menuXML.length => 1 then%>
			<a href="default.asp?id=2" title="Click Here To Switch To The <%=var_usernames_admin%> Login Screen."><%=var_usernames_admin%> Login</a> ::
			<%if var_est_enabled_signup = "1" then%><a href="#" onmouseup="popup('signup.asp?id=1')" title="Click Here To Signup To The System, If You Have Not Already.">Signup</a><%else%>Signup<%end if%> :: 
	<%
	else
	%>		
			<a href="default.asp?id=2" title="Click Here To Switch To The <%=var_usernames_admin%> Login Screen."><%=var_usernames_admin%> Login</a> :: 
			<%if var_est_enabled_signup = "1" then%><a href="#" onmouseup="popup('signup.asp?id=1')" title="Click Here To Signup To The System, If You Have Not Already.">Signup</a><%else%>Signup<%end if%>
	<%
	end if

end if

if menuXML.length = 0 then

else

loopcount = 0
do until loopcount = (menuXML.length - 1)
%>
<a <%if settingsXML.documentElement.childNodes.item(1).childNodes.item(1).childNodes.item(loopcount).getAttribute("href") = "" then%><%else%>href="<%=settingsXML.documentElement.childNodes.item(1).childNodes.item(1).childNodes.item(loopcount).getAttribute("href")%>"<%end if%>title="<%=settingsXML.documentElement.childNodes.item(1).childNodes.item(1).childNodes.item(loopcount).getAttribute("title")%>"><%=settingsXML.documentElement.childNodes.item(1).childNodes.item(1).childNodes.item(loopcount).getAttribute("text")%></a> :: 
<%
loopcount = loopcount + 1
loop
%>
<a <%if settingsXML.documentElement.childNodes.item(1).childNodes.item(1).childNodes.item(loopcount).getAttribute("href") = "" then%><%else%>href="<%=settingsXML.documentElement.childNodes.item(1).childNodes.item(1).childNodes.item(loopcount).getAttribute("href")%>"<%end if%>title="<%=settingsXML.documentElement.childNodes.item(1).childNodes.item(1).childNodes.item(loopcount).getAttribute("title")%>"><%=settingsXML.documentElement.childNodes.item(1).childNodes.item(1).childNodes.item(loopcount).getAttribute("text")%></a>
<%
end if
set menuXML = nothing

end if
%>
		</td>
	</tr>
</table>
</form>

</body>

</html>

<%
set var_comment = nothing
set logintype= nothing
%><!--#include virtual="/pt/modules/ss/p_e.inc"-->