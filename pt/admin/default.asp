<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" >
<!--#include virtual="/pt/modules/ss/usersys/logincheck_admin.inc"-->

<!--#include virtual="/pt/modules/ss/p_s.inc"-->
<%
if session("sess_un") = settingsXML.documentElement.childNodes.item(0).childNodes.item(4).getAttribute("firstloginacc") then
	response.redirect "setup.asp?err=2"
else
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="../modules/css/admin.css">
<script language="javascript" type="text/javascript" src="/pt/modules/js/admin.js"></script>
<title><%=var_ptitle%></title>
</head>

<body>

<div class="smlb_b"></div>
<div class="topb_b"></div>

<div class="main">
	<!--#include virtual="/pt/modules/ss/topbar/admin.inc"-->
<%
if (request("err")) = "1" then
%>
	<div class="m_l">
		<div class="m_l_title">Access Denied</div>
		<div class="m_l_subtitle">You Have Insufficient Privileges To View This Resource!</div>
		<div class="m_l_ins">Please Go Back And Try Another Page.</div>
		<div class="m_l_sel">
		Sorry <%=session("sess_fn")%>, But The Page You Have Tried To Access Has Been Blocked As You Do Not Have The Right Privileges
		To Access It.<br>
		If You Believe This To Be An Error, Please Contact The Administrator Of The System.<br>
		Otherwise, Please Choose An Option From Below To Continue.
			<div class="botopts">
				<ul>
					<li><a href="#" onmouseup="history.back();">Go Back</a></li>
					<li><a href="default.asp?id=1">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
elseif (request("err")) = "2" then
%>
	<div class="m_l">
		<div class="m_l_title">Access Denied</div>
		<div class="m_l_subtitle">You Cannot View This Resource!</div>
		<div class="m_l_ins">Please Go Back And Try Another Page.</div>
		<div class="m_l_sel">
		Sorry <%=session("sess_fn")%>, But Access Has Been Denied To The Setup Wizard.<br>
		This Can Only Be Run Under Certain Conditions, And One Or More Of These Conditions Has Not Been Met.
		<%
		if session("sess_acclevel") <> 1 then
		else
		%>
		<br><br>
		If You Wish To Completley Reset The System And Start Afresh, Please Go To "Settings".
		<%
		end if
		%>
		Please Choose An Option From Below To Continue.
			<div class="botopts">
				<ul>
					<li><a href="#" onmouseup="history.back();">Go Back</a></li>
					<li><a href="default.asp?id=1">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
else
%>
	<div class="m_l">
		<div class="m_l_title"><script language="javascript" type="text/javascript">document.write(daymsg)</script>, <%=session("sess_fn")%>!</div>
		<div class="m_l_subtitle">Welcome To The System!</div>
		<div class="m_l_ins">Please Choose A Task From Below...</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<%
				if session("sess_adminlevel") = "2" then
				%>
				<tr>
					<td class="m_l_sel_t_l"><a href="cover.asp?id=1"><img src="/pt/media/icons/48_wand.png" border="0"></a></td>
					<td class="m_l_sel_t_r" style="font-size: 14pt;"><a href="cover.asp?id=1">Arrange Cover For Your Department</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
				elseif session("sess_adminlevel") = "1" then
				%>
				<tr>
					<td class="m_l_sel_t_l"><a href="cover.asp?id=1"><img src="/pt/media/icons/48_wand.png" border="0"></a></td>
					<td class="m_l_sel_t_r" style="font-size: 14pt;"><a href="cover.asp?id=1">Arrange Cover</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
				else
				end if

				if session("sess_adminlevel") <> 1 then
				%>
				<tr>
					<td class="m_l_sel_t_l"><a href="staff.asp?id=3"><img src="/pt/media/icons/48_man.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="staff.asp?id=3">View Your Staff Timetables</a></td>
				</tr>
				<%
				else
				%>
				<tr>
					<td class="m_l_sel_t_l"><a href="staff.asp?id=1"><img src="/pt/media/icons/48_man.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="staff.asp?id=1">Staff Management</a></td>
				</tr>
				<%
				end if

				if session("sess_adminlevel") <> 1 then
				else
				%>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="rooms.asp?id=1"><img src="/pt/media/icons/48_room.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="rooms.asp?id=1">Room Management</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="ocover.asp?id=1"><img src="/pt/media/icons/48_staffo.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="ocover.asp?id=1">Outside Cover</a></td>
				</tr>
				<%
				end if
				%>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
				if session("sess_adminlevel") <> 1 then
				%>
				<tr>
					<td class="m_l_sel_t_l"><a href="reports.asp?id=3"><img src="/pt/media/icons/48_view.png" border="0" alt="View All Of Today's PleaseTakes In One Table."></a></td>
					<td class="m_l_sel_t_r"><a href="reports.asp?id=3">View Today's Cover Summary</a></td>
				</tr>
				<%
				else
				%>
				<tr>
					<td class="m_l_sel_t_l"><a href="reports.asp?id=1"><img src="/pt/media/icons/48_cal.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="reports.asp?id=1">View Reports</a></td>
				</tr>
				<%
				end if
				%>
				<%
				if session("sess_adminlevel") <> 1 then
				else
				%>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="backup.asp?id=1"><img src="/pt/media/icons/48_bakup.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="backup.asp?id=1">Data Backup</a></td>
				</tr>
				<%
				end if

				if session("sess_adminlevel") <> 1 then
				else
				%>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="settings.asp?id=1"><img src="/pt/media/icons/48_setting.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="settings.asp?id=1">Change Settings</a></td>
				</tr>
				<%
				end if
				%>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="/pt/modules/ss/usersys/logout.asp?id=2"><img src="/pt/media/icons/48_key.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="/pt/modules/ss/usersys/logout.asp?id=2">Logout</a></td>
				</tr>
			</table>
		</div>
	</div>
<%
end if
%>
	<div class="m_r">
		<div class="m_r2">
		<!--#include virtual="/pt/modules/ss/rbar/admin.inc"-->
		</div>
	</div>
</div>

</body>

</html>
<%
end if
%><!--#include virtual="/pt/modules/ss/p_e.inc"-->