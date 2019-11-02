<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" >
<!--#include virtual="/pt/modules/ss/usersys/logincheck_admin.inc"-->
<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->

<!--#include virtual="/pt/modules/ss/p_s.inc"-->
<%
if session("sess_un") = settingsXML.documentElement.childNodes.item(0).childNodes.item(4).getAttribute("firstloginacc") then
	response.redirect "setup.asp?err=2"
else

pagetype = request("id")
backuptype = request("backup")

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
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="../modules/css/admin.css">
<%
if pagetype = "3" then
%>
<link rel="alternate" media="print" href="backup.asp?id=3&amp;print=1&amp;backup=<%=request("backup")%>">
<%
else
end if
%>
<script language="javascript" type="text/javascript" src="/pt/modules/js/admin.js"></script>
<%if (pagetype = "2") AND (backuptype = "2") then%>
<script language="javascript" type="text/javascript">
function compredir()
{
 location.href = 'backup.asp?id=2&backup=3';
}
</script>
<%
else
end if
%>
<title><%=var_ptitle%></title>
</head>

<body<%if (pagetype = "2") AND (backuptype = "2") then%> onLoad="setTimeout('compredir()', 7000)"<%else%><%end if%>>

<div class="smlb_b"></div>
<div class="topb_b"></div>

<div class="main">
	<!--#include virtual="/pt/modules/ss/topbar/admin.inc"-->

<%
if pagetype = "2" then

	if backuptype = "2" then
%>
	<!--#include virtual="/pt/modules/ss/alerts/admin_onlylday.inc"-->
<%
	coverbackupSQL = "SELECT * INTO [C_" & getweekstart(date()) & "_" & getweekend(date()) & "_" & getweekno(date()) & "] IN 'D:\Websites\Data\PleaseTakes\backup.mdb' FROM Cover WHERE (((Cover.DAYDATE) Between #" & SQLDate(getweekstart(date())) & "# And #" & SQLDate(getweekend(date())) & "#))"
	dataconn.Execute(coverbackupSQL)
	attbackupSQL = "SELECT * INTO [A_" & getweekstart(date()) & "_" & getweekend(date()) & "_" & getweekno(date()) & "] IN 'D:\Websites\Data\PleaseTakes\backup.mdb' FROM Attendance WHERE (((Attendance.DAYDATE) Between #" & SQLDate(getweekstart(date())) & "# And #" & SQLDate(getweekend(date())) & "#))"
	dataconn.Execute(attbackupSQL)
	backupinSQL = "INSERT INTO Inventory ([StartDate],[EndDate],[WeekNo],[BackupMonth],[BackupYear]) Values ('" & GetWeekStart(date()) & "','" & GetWeekEnd(date()) & "','" & GetWeekNo(date()) & "','" & month(date()) & "','" & year(date()) & "')"
	backupconn.Execute(backupinSQL)

	coverclrSQL = "DELETE * FROM Cover WHERE (((Cover.DAYDATE) Between #" & SQLDate(getweekstart(date())) & "# And #" & SQLDate(getweekend(date())) & "#))"
	dataconn.Execute(coverclrSQL)
	attclrSQL = "DELETE * FROM Attendance WHERE (((Attendance.DAYDATE) Between #" & SQLDate(getweekstart(date())) & "# And #" & SQLDate(getweekend(date())) & "#))"
	dataconn.Execute(attclrSQL)
%>
	<div class="m_l">
		<div class="m_l_title">Please Wait...</div>
		<div class="m_l_subtitle">Backing Up...</div>
		<div class="m_l_ins">Please DO NOT Close This Window!</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<span style="font-size: 12pt; font-weight: bold; letter-spacing: -1px;">
			Please Note That The Data For This Week Has Been Moved To A Special "Backup" Database, So It Appears That All The Data From This Week Has Disappered. This Is NOT The Case.
			</span>
			Please Wait While The System Backs Up All Data To The Backup Database.<br>
			This Will Take About Seven Seconds To Complete, After Which Point You Will Be Taken To The Completion Page.<br>
			<b>Please DO NOT Close This Window!</b><br>
			<div style="text-align: center;"><img src="/pt/media/admin/progress.gif" border="0" alt="Progress Boxes, A Fresh Alternative From A Bar."></div>
		</div>
	</div>
<%
	elseif backuptype = "3" then
%>
	<div class="m_l">
		<div class="m_l_title">Backup Complete</div>
		<div class="m_l_subtitle">All Data Has Been Backed Up Successfully</div>
		<div class="m_l_ins">Congratulations!</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Congratulations, <%=session("sess_fn")%>! You Have Successfully Backed Up All The Data Of This Week Into The Backup Database.
			The Master Database Has Now Been Cleared And Is Ready For Next Week's Cover Information.<br>
			<b>This Task Will Be Unavailable Until Next <%if var_est_enabled_weekends = 1 then%>Saturday<%else%>Friday<%end if%> Afternoon.</b>
			<div class="botopts">
				<ul>
					<li><a href="backup.asp?id=1">Return To The Data Backup Homepage</a></li>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	else
%>
	<!--#include virtual="/pt/modules/ss/alerts/admin_onlylday.inc"-->
	<div class="m_l">
		<div class="m_l_title">Backup Weekly Data</div>
		<div class="m_l_subtitle">Backup All Of The Week's Data, Leaving You Ready For The Next.</div>
		<div class="m_l_ins">Have A Read At What Is Below, And Then Click "Begin".</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">Running This Task Before <%if var_est_enabled_weekends = 1 then%>Saturday<%else%>Friday<%end if%> Afternoon/Evening Will Cause Problems!</td>
				</tr>
			</table>
			<hr size="1">
			This One-Step Process Will Backup All Of The Data Gathered Over The Week And Place It In Another Database.
			This Database Is Designated The "Backup Database", And Is Where All Old Data Is Stored.<br>
			<p><b>To Backup All The Weekly Data, Please Click "Begin". This Process Is Irreversible.</b>
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="backup.asp?id=2&amp;backup=2">Begin></a></div>
		</div>
	</div>
<%
	end if

elseif pagetype = "3" then
	if backuptype = "" then
		response.redirect "/pt/admin/backup.asp?id=1"
	else
		if (request("print") = "1") then
			response.redirect "/pt/admin/print.asp?id=5&backup=" & backuptype
		else
	
		backupyear = left(backuptype,4)
		backupweek = mid(backuptype,6)

		RSBACKEDUPSQL = "SELECT * FROM Inventory WHERE weekno = " & backupweek & " AND backupyear = " & backupyear

		Set RSBACKEDUP = Server.CreateObject("Adodb.RecordSet")
		RSBACKEDUP.Open RSBACKEDUPSQL, backupconn, adopenkeyset, adlockoptimistic
		
			if RSBACKEDUP.RECORDCOUNT = 0 then
				response.redirect "backup.asp?id=1"
			else

			RSCOVERSQL = "SELECT * FROM [C_" & RSBACKEDUP("StartDate") & "_" & RSBACKEDUP("EndDate") & "_" & RSBACKEDUP("WeekNo") & "]"

			Set RSCOVER = Server.CreateObject("Adodb.RecordSet")
			RSCOVER.Open RSCOVERSQL, backupconn, adopenkeyset, adlockoptimistic

			RSATTSQL = "SELECT * FROM [A_" & RSBACKEDUP("StartDate") & "_" & RSBACKEDUP("EndDate") & "_" & RSBACKEDUP("WeekNo") & "]"

			Set RSATT = Server.CreateObject("Adodb.RecordSet")
			RSATT.Open RSATTSQL, backupconn, adopenkeyset, adlockoptimistic
%>
	<div class="m_l">
		<div class="m_l_title">Data For Week <%=RSBACKEDUP("StartDate")%> - <%=RSBACKEDUP("EndDate")%></div>
		<div class="m_l_subtitle">See Who Was Absent, And Who Covered Them Over The Week.</div>
		<div class="m_l_ins">For A Printer-Friendly Version, Please Click <a href="#" onmouseup="location.href='backup.asp?id=3&amp;print=1&amp;backup=<%=backuptype%>'"><img src="/pt/media/icons/16_print.gif" border="0" alt="Click Here To Print The Report For Backed Up Data.">&nbsp;Here</a>.</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
		Coming Soon...
		</div>
	</div>
<%
			end if
		end if
	end if
else
%>
	<div class="m_l">
		<div class="m_l_title">Data Backup</div>
		<div class="m_l_subtitle">View Backed Up Data Or Backup This Week's Data.</div>
		<div class="m_l_ins">Have A Read At What Is Below, And Then Click "Next".</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<%
			if (request("err") = "1") then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t" style="color: #F00;">The Data Backup Tool Cannot Be Used Until <%if var_est_enabled_weekends = 1 then%>Saturday<%else%>Friday<%end if%>  After <%=var_backuptime%>:00, To Avoid Problems!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			else
			end if
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_t_l"><a href="backup.asp?id=2"><img src="/pt/media/icons/48_wand.png" border="0" alt="Use The Data Backup Wizard To Backup The Week's Data, So You Can Be Prepared For The Next Week."></a></td>
					<td class="m_l_sel_t_r" style="font-size: 14pt;"><a href="backup.asp?id=2">Backup This Week's Data</a></td>
				
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
			</table>
			<%
			RSBACKEDUPSQL = "SELECT * FROM Inventory ORDER BY StartDate DESC"
		
			Set RSBACKEDUP = Server.CreateObject("Adodb.RecordSet")
			RSBACKEDUP.Open RSBACKEDUPSQL, backupconn, adopenkeyset, adlockoptimistic
			
			if RSBACKEDUP.RECORDCOUNT = 0 then
			%>
			<div class="m_l_middletitle" style="width: 546px; _width: 100%; margin-bottom: 10px; font-size: 12pt;">No Backed Up Data Found!</div>
			<%
			else
			%>
			<div class="m_l_middletitle" style="width: 546px; _width: 100%; margin-bottom: 10px; font-size: 12pt;">View Previously Backed Up Data By Selecting A Week From Below....</div>
			<div style="width: 540px; _width: 100%; padding: 0px 0px 5px 10px; margin-top: -10px; line-height: 28px;">
				<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
					<tr>
						<td class="m_l_sel_sep" colspan="2"></td>
					</tr>
				<%
				do until RSBACKEDUP.EOF
					RSQSTATSSQL = "SELECT * FROM [C_" & RSBACKEDUP("StartDate") & "_" & RSBACKEDUP("EndDate") & "_" & RSBACKEDUP("WeekNo") & "]"
		
					Set RSQSTATS = Server.CreateObject("Adodb.RecordSet")
					RSQSTATS.Open RSQSTATSSQL, backupconn, adopenkeyset, adlockoptimistic

					RSQSTATSISQL = "SELECT * FROM [C_" & RSBACKEDUP("StartDate") & "_" & RSBACKEDUP("EndDate") & "_" & RSBACKEDUP("WeekNo") & "] WHERE OCOVER <> 1"
		
					Set RSQSTATSI = Server.CreateObject("Adodb.RecordSet")
					RSQSTATSI.Open RSQSTATSISQL, backupconn, adopenkeyset, adlockoptimistic
				%>
					<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('<%=RSBACKEDUP("backupyear")%>_<%=RSBACKEDUP("weekno")%>');">
						<td class="m_l_list_p"><img src="/pt/media/icons/16_backup.gif" border="0" alt="View Backed Up Data For The Week Beginning <%=RSBACKEDUP("StartDate")%>."></td>
						<td class="m_l_list_t">Data For Week <b><%=RSBACKEDUP("StartDate")%> - <%=RSBACKEDUP("EndDate")%></b></td>
					</tr>
					<tr id="list_<%=RSBACKEDUP("backupyear")%>_<%=RSBACKEDUP("weekno")%>" style="display: none;">
						<td class="m_l_list_p"></td>
						<td class="m_l_list_t" style="line-height: 20px; padding-top: 7px;">
						<b>Total Cover Requests:</b> <%=RSQSTATS.RECORDCOUNT%><br>
						<b>Staff Staff ("Inside") Cover:</b> <%=RSQSTATSI.RECORDCOUNT%><br>
						<b>Total Outside Cover:</b> <%=RSQSTATS.RECORDCOUNT - RSQSTATSI.RECORDCOUNT%>
						<hr size="1">
						<div style="width: 100%; text-align: center;"><b><a href="backup.asp?id=3&amp;backup=<%=RSBACKEDUP("backupyear")%>_<%=RSBACKEDUP("weekno")%>">View Week's Data</a></b></div>						
						</td>
					</tr>
					<tr>
						<td class="m_l_sel_sep" colspan="2"></td>
					</tr>
				<%
					RSQSTATS.close
					RSQSTATSI.close
					set RSQSTATS = nothing
					set RSQSTATSI = nothing
				RSBACKEDUP.MOVENEXT
				loop
				%>
				</table>
			</div>
			<%
			end if
			%>
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