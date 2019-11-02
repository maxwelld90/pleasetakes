<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" >
<!--#include virtual="/pt/modules/ss/usersys/logincheck_admin.inc"-->

<!--#include virtual="/pt/modules/ss/p_s.inc"-->
<%
if session("sess_un") = settingsXML.documentElement.childNodes.item(0).childNodes.item(4).getAttribute("firstloginacc") then
	response.redirect "setup.asp?err=2"
else

pagetype = request("id")
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="../modules/css/admin.css">
<%
if pagetype = "2" then
%>
<link rel="alternate" media="print" href="reports.asp?id=2&amp;print=1">
<%
elseif pagetype = "3" then
%>
<link rel="alternate" media="print" href="reports.asp?id=3&amp;print=1">
<%
elseif pagetype = "4" then
%>
<link rel="alternate" media="print" href="reports.asp?id=4&amp;print=1">
<%
else
end if
%>
<script language="javascript" type="text/javascript" src="/pt/modules/js/admin.js"></script>
<title><%=var_ptitle%></title>
</head>

<body>

<div class="smlb_b"></div>
<div class="topb_b"></div>

<div class="main">
	<!--#include virtual="/pt/modules/ss/topbar/admin.inc"-->

<%
if pagetype = "2" then

	noslipsreq = 0

	if (request("dow") = "") then
		daydow = dow
	else
		daydow = request("dow")
	end if
	if (request("daydate") = "") then
		daydate = date()
	else
		daydate = request("daydate")
	end if

	if (request("print")) = "1" then
		response.redirect "/pt/admin/print.asp?id=2&dow=" & daydow & "&daydate=" & daydate
	else
%>
	<div class="m_l">
		<div class="m_l_title">View/Print PleaseTake Slips</div>
		<div class="m_l_subtitle">View Or Print Slips To Be Handed Out To Staff Members.</div>
		<div class="m_l_ins">For A Printer Friendly Version, Please Click <a href="#" onmouseup="location.href='reports.asp?id=2&amp;print=1&amp;dow=<%=daydow%>&amp;daydate=<%=daydate%>'"><img src="/pt/media/icons/16_print.gif" border="0" alt="Click Here To Print Today's PleaseTake Slips.">&nbsp;Here</a>.</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<form name="edit" action="/pt/modules/ss/db/edit.asp?edittype=15" method="post">
				<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
					<tr>
						<td class="m_l_list_t" style="width: 85px;">Select A Day:</td>
						<td class="m_l_list_t" style="width: 180px;">
							<select class="b_std" size="1" name="REP3_CHOICE" onchange="document.edit.submit();">
								<option value="">Please Select...</option>
								<%
								if var_est_enabled_weekends = "1" then
									for i = 1 to 7
								%>
								<option value="<%=i%>,<%=DateAdd("d",i - 1,getweekstart(date()))%>"><%=getdow((DateAdd("d",i - 1,getweekstart(date()))))%>, <%=DateAdd("d",i - 1,getweekstart(date()))%></option>
								<%
									next
								else
									for i = 2 to 6
								%>
								<option value="<%=i%>,<%=DateAdd("d",i - 1,getweekstart(date()))%>"><%=getdow((DateAdd("d",i - 1,getweekstart(date()))))%>, <%=DateAdd("d",i - 1,getweekstart(date()))%></option>
								<%	
									next
								end if
								%>
							</select>
						</td>
						<td class="m_l_list_t"><b>Currently Viewing</b> <%if (daydow = dow) and (cdate(daydate) = date()) then%>Today<%elseif (request("dow") = "") and (request("daydate") = "") then%>Today<%else%><%=getdow(daydow)%>, <%=daydate%><%end if%></td>
					</tr>
				</table>
			</form>
			<hr size="1">
			<div class="m_l_middletitle" style="width: 546px; _width: 100%; margin-bottom: 10px; font-size: 12pt;">Staff Cover</div>	
			<%
			RSSLIPSSQL = "SELECT * FROM Cover WHERE DAY = " & daydow & " AND DAYDATE = #" & SQLDate(daydate) & "# AND OCOVER <> 1 ORDER BY PERIOD"
		
			Set RSSLIPS = Server.CreateObject("Adodb.RecordSet")
			RSSLIPS.Open RSSLIPSSQL, dataconn, adopenkeyset, adlockoptimistic
			
			if RSSLIPS.RECORDCOUNT = 0 then
				RSSLIPS.close
				set RSSLIPS = nothing
				noslipsreq = 1
			%>
			<div style="padding: 5px 0px 5px 10px; margin-top: -10px; line-height: 28px;">
			There Are No Staff Cover Slips To Be Printed For <%=getdow(daydow)%>, <%=daydate%>!
			</div>
			<%
			else

			do until RSSLIPS.EOF
			
				RSCOVERSQL = "SELECT * FROM Cover WHERE FOR = " & RSSLIPS("FOR") & " AND DAY = " & daydow & " AND DAYDATE = #" & SQLDate(daydate) & "#"
				Set RSCOVER = Server.CreateObject("Adodb.RecordSet")
				RSCOVER.Open RSCOVERSQL, dataconn, adopenkeyset, adlockoptimistic

				RSCLASSSQL = "SELECT [" & RSSLIPS("PERIOD") & "_" & daydow & "], [R" & RSSLIPS("PERIOD") & "_" & daydow & "] FROM Timetables WHERE ID = " & RSSLIPS("FOR")	
				Set RSCLASS = Server.CreateObject("Adodb.RecordSet")
				RSCLASS.Open RSCLASSSQL, dataconn, adopenkeyset, adlockoptimistic

				RSFORSQL = "SELECT * FROM Timetables WHERE ID = " & RSSLIPS("FOR")	
				Set RSFOR = Server.CreateObject("Adodb.RecordSet")
				RSFOR.Open RSFORSQL, dataconn, adopenkeyset, adlockoptimistic

					RSFORDEPTSQL = "SELECT * FROM Departments WHERE DEPTID = " & RSFOR("DEPT")		
					Set RSFORDEPT = Server.CreateObject("Adodb.RecordSet")
					RSFORDEPT.Open RSFORDEPTSQL, dataconn, adopenkeyset, adlockoptimistic		
		
				RSCOVERINGSQL = "SELECT * FROM Timetables WHERE ID = " & RSSLIPS("COVERING")
				Set RSCOVERING = Server.CreateObject("Adodb.RecordSet")
				RSCOVERING.Open RSCOVERINGSQL, dataconn, adopenkeyset, adlockoptimistic
	
					RSCOVERINGDEPTSQL = "SELECT * FROM Departments WHERE DEPTID = " & RSCOVERING("DEPT")
					Set RSCOVERINGDEPT = Server.CreateObject("Adodb.RecordSet")
					RSCOVERINGDEPT.Open RSCOVERINGDEPTSQL, dataconn, adopenkeyset, adlockoptimistic	
			%>
			<div class="rep_slip">
				<div class="rep_slip_c">
					<div style="height: 22px; line-height: 15px; border-bottom: 1px solid #D99251;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr>
								<td style="width: 50%; font-size: 10pt; font-family: Tahoma,Sans-Serif;"><%=getdow(daydow)%>, <%=day(daydate)%>&nbsp;<%=cal_getmonth(month(daydate))%>&nbsp;<%=year(daydate)%></td>
								<td style="width: 50%; font-size: 10pt; text-align: right; font-family: Tahoma,Sans-Serif;"><b><%=var_pname%></b> <%=var_ver%></td>
							</tr>
						</table>
					</div>
					<div class="rep_slip_b">
					<table class="m_l_sel_t" style="height: 100%;" cellpadding="0" cellspacing="0">
						<tr>
							<td style="width: 60%;">
							<table class="m_l_sel_t" style="height: 100%;" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="2" class="rep_slip_t" style="border-bottom: 1px solid #D99251;"><b>Please Arrange To Take This Class.</b></td>
								</tr>
								<tr>
									<td class="rep_slip_t" style="height: 30px; width: 25%; border-bottom: 1px solid #D99251;">&nbsp;<b>Covering:&nbsp;&nbsp;</b></td>
									<td class="rep_slip_t" style="border-bottom: 1px solid #D99251;"><b><%=left(RSCOVERING("FN"),1)%>.&nbsp;<%=RSCOVERING("LN")%></b> (<%=RSCOVERINGDEPT("SHORT")%><%if RSCOVERING("DEFROOM") <> "" then%>, Room <b><%=RSCOVERING("DEFROOM")%></b><%else%><%end if%>)</td>
								</tr>
								<tr>
									<td class="rep_slip_t" style="height: 30px; width: 25%; border-bottom: 1px solid #D99251;">&nbsp;<b>For:</b></td>
									<td class="rep_slip_t" style="border-bottom: 1px solid #D99251;"><%=left(RSFOR("FN"),1)%>.&nbsp;<%=RSFOR("LN")%> (<%=RSFORDEPT("SHORT")%>)</td>
								</tr>
								<tr>
									<td class="rep_slip_t" style="height: 30px; width: 25%; border-bottom: 1px solid #D99251;">&nbsp;<b>Class:</b></td>
									<td class="rep_slip_t" style="border-bottom: 1px solid #D99251;"><%=RSCLASS(RSSLIPS("PERIOD") & "_" & daydow)%></td>
								</tr>
								<tr>
									<td colspan="2" class="rep_slip_t" style="text-align: right;"><b>Thank You Very Much!</b>&nbsp;</td>
								</tr>
							</table>
							</td>
							<td style="width: 40%; text-align: right;">
							<table class="m_l_sel_t" style="height: 100%;" cellpadding="0" cellspacing="0">
								<tr>
									<td class="rep_slip_t" style="width: 50%; height: 27px; border-bottom: 1px solid #D99251; border-left: 1px solid #D99251; border-right: 1px solid #D99251; text-align: center;"><b>Period</b></td>
									<td class="rep_slip_t" style="width: 50%; height: 27px; border-bottom: 1px solid #D99251; text-align: center;"><b>Room</b></td>
								</tr>
								<tr>
									<td class="rep_slip_t" style="border-left: 1px solid #D99251; border-right: 1px solid #D99251; text-align: center; font-size: 70pt;"><%=RSSLIPS("PERIOD")%></td>
									<td class="rep_slip_t" style="text-align: center; font-size: 36pt; letter-spacing: -6px;"><%=RSCLASS("R" & RSSLIPS("PERIOD") & "_" & daydow)%></td>
								</tr>
							</table>
							</td>
						</tr>
					</table>
					</div>
				</div>
			</div>
			<div class="m_l_sel_sep"></div>
			<%
				RSCOVERINGDEPT.close
				RSCOVERING.close
				RSFORDEPT.close
				RSFOR.close
				RSCLASS.close
				RSCOVER.close
				set RSCOVERINGDEPT = nothing
				set RSCOVERING = nothing
				set RSFORDEPT = nothing
				set RSFOR = nothing
				set RSCLASS = nothing
				set RSCOVER = nothing

			RSSLIPS.MOVENEXT
			loop

			RSSLIPS.close
			set RSSLIPS = nothing

			end if
			%>
			<div class="m_l_middletitle" style="width: 546px; _width: 100%; margin-bottom: 10px; font-size: 12pt;">Outside Cover</div>
			<%
			RSOCOVERNAMESSQL = "SELECT * FROM OCover ORDER BY LN ASC"
		
			Set RSOCOVERNAMES = Server.CreateObject("Adodb.RecordSet")
			RSOCOVERNAMES.Open RSOCOVERNAMESSQL, dataconn, adopenkeyset, adlockoptimistic
		
			RSOSLIPSSQL = "SELECT * FROM Cover INNER JOIN OCover ON Cover.COVERING = OCover.ID WHERE (((Cover.DAY)=" & daydow & ") AND ((Cover.DAYDATE)=#" & SQLDate(daydate) & "#) AND ((Cover.OCOVER)=1)) ORDER BY OCover.LN, Cover.Period"

			Set RSOSLIPS = Server.CreateObject("Adodb.RecordSet")
			RSOSLIPS.Open RSOSLIPSSQL, dataconn, adopenkeyset, adlockoptimistic
			
			if RSOSLIPS.RECORDCOUNT = 0 then
				RSOSLIPS.close
				set RSOSLIPS = nothing
				noslipsreq = 1
			%>
			<div style="padding: 5px 0px 10px 10px; margin-top: -10px; line-height: 28px;">
			No Outside Cover Slips Are To Be Printed For <%=getdow(daydow)%>, <%=daydate%>!
			</div>
			<%
			else

			do until RSOCOVERNAMES.EOF
			do until RSOSLIPS.EOF
			
				RSCOVERSQL = "SELECT * FROM Cover WHERE FOR = " & RSOSLIPS("FOR") & " AND DAY = " & daydow & " AND DAYDATE = #" & SQLDate(daydate) & "#"
				Set RSCOVER = Server.CreateObject("Adodb.RecordSet")
				RSCOVER.Open RSCOVERSQL, dataconn, adopenkeyset, adlockoptimistic

				RSCLASSSQL = "SELECT [" & RSOSLIPS("PERIOD") & "_" & daydow & "], [R" & RSOSLIPS("PERIOD") & "_" & daydow & "] FROM Timetables WHERE ID = " & RSOSLIPS("FOR")	
				Set RSCLASS = Server.CreateObject("Adodb.RecordSet")
				RSCLASS.Open RSCLASSSQL, dataconn, adopenkeyset, adlockoptimistic

				RSFORSQL = "SELECT * FROM Timetables WHERE ID = " & RSOSLIPS("FOR")	
				Set RSFOR = Server.CreateObject("Adodb.RecordSet")
				RSFOR.Open RSFORSQL, dataconn, adopenkeyset, adlockoptimistic

					RSFORDEPTSQL = "SELECT * FROM Departments WHERE DEPTID = " & RSFOR("DEPT")		
					Set RSFORDEPT = Server.CreateObject("Adodb.RecordSet")
					RSFORDEPT.Open RSFORDEPTSQL, dataconn, adopenkeyset, adlockoptimistic		
		
				RSCOVERINGSQL = "SELECT * FROM OCover WHERE ID = " & RSOSLIPS("COVERING")
				Set RSCOVERING = Server.CreateObject("Adodb.RecordSet")
				RSCOVERING.Open RSCOVERINGSQL, dataconn, adopenkeyset, adlockoptimistic
			%>
			<div class="rep_slip">
				<div class="rep_slip_c">
					<div style="height: 22px; line-height: 15px; border-bottom: 1px solid #D99251;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr>
								<td style="width: 50%; font-size: 10pt; font-family: Tahoma,Sans-Serif;"><%=getdow(daydow)%>, <%=day(daydate)%>&nbsp;<%=cal_getmonth(month(daydate))%>&nbsp;<%=year(daydate)%></td>
								<td style="width: 50%; font-size: 10pt; text-align: right; font-family: Tahoma,Sans-Serif;"><b><%=var_pname%></b> <%=var_ver%></td>
							</tr>
						</table>
					</div>
					<div class="rep_slip_b">
					<table class="m_l_sel_t" style="height: 100%;" cellpadding="0" cellspacing="0">
						<tr>
							<td style="width: 60%;">
							<table class="m_l_sel_t" style="height: 100%;" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="2" class="rep_slip_t" style="border-bottom: 1px solid #D99251;"><b>Please Arrange To Take This Class.</b></td>
								</tr>
								<tr>
									<td class="rep_slip_t" style="height: 30px; width: 25%; border-bottom: 1px solid #D99251;">&nbsp;<b>Covering:&nbsp;&nbsp;</b></td>
									<td class="rep_slip_t" style="border-bottom: 1px solid #D99251;"><b><%=left(RSCOVERING("FN"),1)%>.&nbsp;<%=RSCOVERING("LN")%></b> (Outside Cover)</td>
								</tr>
								<tr>
									<td class="rep_slip_t" style="height: 30px; width: 25%; border-bottom: 1px solid #D99251;">&nbsp;<b>For:</b></td>
									<td class="rep_slip_t" style="border-bottom: 1px solid #D99251;"><%=left(RSFOR("FN"),1)%>.&nbsp;<%=RSFOR("LN")%> (<%=RSFORDEPT("SHORT")%>)</td>
								</tr>
								<tr>
									<td class="rep_slip_t" style="height: 30px; width: 25%; border-bottom: 1px solid #D99251;">&nbsp;<b>Class:</b></td>
									<td class="rep_slip_t" style="border-bottom: 1px solid #D99251;"><%=RSCLASS(RSOSLIPS("PERIOD") & "_" & daydow)%></td>
								</tr>
								<tr>
									<td colspan="2" class="rep_slip_t" style="text-align: right;"><b>Thank You Very Much!</b>&nbsp;</td>
								</tr>
							</table>
							</td>
							<td style="width: 40%; text-align: right;">
							<table class="m_l_sel_t" style="height: 100%;" cellpadding="0" cellspacing="0">
								<tr>
									<td class="rep_slip_t" style="width: 50%; height: 27px; border-bottom: 1px solid #D99251; border-left: 1px solid #D99251; border-right: 1px solid #D99251; text-align: center;"><b>Period</b></td>
									<td class="rep_slip_t" style="width: 50%; height: 27px; border-bottom: 1px solid #D99251; text-align: center;"><b>Room</b></td>
								</tr>
								<tr>
									<td class="rep_slip_t" style="border-left: 1px solid #D99251; border-right: 1px solid #D99251; text-align: center; font-size: 70pt;"><%=RSOSLIPS("PERIOD")%></td>
									<td class="rep_slip_t" style="text-align: center; font-size: 36pt; letter-spacing: -6px;"><%=RSCLASS("R" & RSOSLIPS("PERIOD") & "_" & daydow)%></td>
								</tr>
							</table>
							</td>
						</tr>
					</table>
					</div>
				</div>
			</div>
			<div class="m_l_sel_sep"></div>
			<%
				RSCOVERING.close
				RSFORDEPT.close
				RSFOR.close
				RSCLASS.close
				RSCOVER.close
				set RSCOVERING = nothing
				set RSFORDEPT = nothing
				set RSFOR = nothing
				set RSCLASS = nothing
				set RSCOVER = nothing

			RSOSLIPS.MOVENEXT
			loop

			RSOCOVERNAMES.MOVENEXT
			loop
			RSOSLIPS.close
			set RSOSLIPS = nothing
			end if
			if noslipsreq = 1 then
			%>
				<div class="botopts">
					<ul>
						<li><a href="cover.asp?id=1">Arrange Staff Cover</a></li>
						<li><a href="ocover.asp?id=1">Arrange Outside Cover</a></li>
						<li><a href="reports.asp?id=1">Back To The Reports Homepage</a></li>
						<li><a href="default.asp?id=1">Back To The Admin Homepage</a></li>
					</ul>
				</div>
			<%
			else
			end if
			%>
		</div>
	</div>
<%
	end if

elseif pagetype = "3" then

	if (request("dow") = "") then
		daydow = dow
	else
		daydow = request("dow")
	end if
	if (request("daydate") = "") then
		daydate = date()
	else
		daydate = request("daydate")
	end if

	if (request("print")) = "1" then
		response.redirect "/pt/admin/print.asp?id=3&dow=" & daydow & "&daydate=" & daydate
	else
%>
	<div class="m_l">
		<div class="m_l_title">Cover Summary</div>
		<div class="m_l_subtitle">View All Of The PleaseTakes For The Chosen Day In One Table.</div>
		<div class="m_l_ins">For A Printer Friendly Version, Please Click <a href="#" onmouseup="location.href='reports.asp?id=3&amp;print=1&amp;dow=<%=daydow%>&amp;daydate=<%=daydate%>'"><img src="/pt/media/icons/16_print.gif" border="0" alt="Click Here To Print Today's Cover Summary Timetable.">&nbsp;Here</a>.</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<form name="edit" action="/pt/modules/ss/db/edit.asp?edittype=12" method="post">
				<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
					<tr>
						<td class="m_l_list_t" style="width: 85px;">Select A Day:</td>
						<td class="m_l_list_t" style="width: 180px;">
							<select class="b_std" size="1" name="REP3_CHOICE" onchange="document.edit.submit();">
								<option value="">Please Select...</option>
								<%
								if var_est_enabled_weekends = "1" then
									for i = 1 to 7
								%>
								<option value="<%=i%>,<%=DateAdd("d",i - 1,getweekstart(date()))%>"><%=getdow((DateAdd("d",i - 1,getweekstart(date()))))%>, <%=DateAdd("d",i - 1,getweekstart(date()))%></option>
								<%
									next
								else
									for i = 2 to 6
								%>
								<option value="<%=i%>,<%=DateAdd("d",i - 1,getweekstart(date()))%>"><%=getdow((DateAdd("d",i - 1,getweekstart(date()))))%>, <%=DateAdd("d",i - 1,getweekstart(date()))%></option>
								<%	
									next
								end if
								%>
							</select>
						</td>
						<td class="m_l_list_t"><b>Currently Viewing</b> <%if (daydow = dow) and (cdate(daydate) = date()) then%>Today<%elseif (request("dow") = "") and (request("daydate") = "") then%>Today<%else%><%=getdow(daydow)%>, <%=daydate%><%end if%></td>
					</tr>
				</table>
			</form>
			<hr size="1">
			<div style="padding-bottom: 5px;">
			To View More Information And Options On What You Can Do For A Cover Request, Just Click Its Box, And A Popup Window Will Appear.
			</div>
			<!--#include virtual="/pt/modules/ss/timetables/admin_rep_sum.inc"-->
			<%
			RSDAYSQL = "SELECT * FROM Periods WHERE ID = " & daydow

			Set RSDAY = Server.CreateObject("Adodb.RecordSet")
			RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic

			RSTOTABSENTSQL = "SELECT * FROM Attendance WHERE DAY = " & daydow & " AND DAYDATE = # " & SQLDate(daydate) & " #"

			Set RSTOTABSENT = Server.CreateObject("Adodb.RecordSet")
			RSTOTABSENT.Open RSTOTABSENTSQL, dataconn, adopenkeyset, adlockoptimistic
			
			if RSTOTABSENT.RECORDCOUNT = 0 then
			else
			%>
			<div class="m_l_middletitle" style="width: 546px; _width: 100%; margin-top: 10px; margin-bottom: 10px; font-size: 12pt;">Problems With <%if (daydow = dow) and (cdate(daydate) = date()) then%>Today's<%elseif (request("dow") = "") and (request("daydate") = "") then%>Today's<%else%><%=getdow(daydow)%>'s<%end if%> Cover</div>
			<div style="padding: 5px 0px 10px 10px; margin-top: -10px; line-height: 28px;">
			<%
			errorless = 0
			
			do until RSTOTABSENT.EOF
				RSUSERSQL = "SELECT * FROM Timetables WHERE ID = " & RSTOTABSENT("USER")

				Set RSUSER = Server.CreateObject("Adodb.RecordSet")
				RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
				
				for i=1 to RSDAY("Totals")
					RSCOVERNSQL = "SELECT * FROM Cover WHERE FOR = " & RSTOTABSENT("USER") & " AND DAY = " & daydow & " AND DAYDATE = # " & SQLDate(daydate) & " # AND PERIOD = " & i

					Set RSCOVERN = Server.CreateObject("Adodb.RecordSet")
					RSCOVERN.Open RSCOVERNSQL, dataconn, adopenkeyset, adlockoptimistic
					if (RSUSER(i & "_" & daydow) <> "") and (RSTOTABSENT(i & "_" & daydow) <> "") and (RSCOVERN.RECORDCOUNT = 0) then
					%>
					No Cover Selected For <span style="font-size: 14pt; letter-spacing: -1px;"><b><%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%></b></span>, Period <b><%=i%></b><br>
					<%
					else
						errorless = errorless + 1
					end if
				next
			RSTOTABSENT.MOVENEXT
			loop

			if errorless = (RSDAY("TOTALS") * RSTOTABSENT.RECORDCOUNT) then
			%>
			<b>No Problems Have Been Found With The Cover Above.</b>
			<%
			else
			end if
			%>
			</div>
			<%
			end if
			%>
		</div>
	</div>
<%
	end if

elseif pagetype = "4" then

	if (request("print")) = "1" then
		response.redirect "/pt/admin/print.asp?id=4"
	else
%>
	<div class="m_l">
		<div class="m_l_title">The Weekly Report</div>
		<div class="m_l_subtitle">The Report That Summarises Everything Over The Past Week.</div>
		<div class="m_l_ins">For A Printer Friendly Version, Please Click <a href="#" onmouseup="printpage(4);"><img src="/pt/media/icons/16_print.gif" border="0" alt="Click Here To Print The Weekly Report.">&nbsp;Here</a>.</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Please Note, <b>Don't</b> Print <b>This</b> Page! A Much Nicer, Printer-Friendly Version Is Available By Clicking The Printer Icon Above.
			<hr size="1">
			<div class="m_l_middletitle" style="width: 544px; _width: 550px; margin-bottom: 10px;" onmouseup="showdetail('contents');">Report Contents</div>
			<div id="list_contents" style="padding: 0px 0px 10px 4px;">
				<table class="m_l_sel_t" style="width: 544px; _width: 546px;" cellpadding="0" cellspacing="0">
					<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('absent');">
						<td class="m_l_list_p m_l_list_t"><b>1</b></td>
						<td class="m_l_list_t">Who Was Absent During The Week?</td>
					</tr>
					<tr>
						<td class="m_l_sel_sep" colspan="2"></td>
					</tr>
					<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('dept');">
						<td class="m_l_list_p m_l_list_t"><b>2</b></td>
						<td class="m_l_list_t">Departmental Statistics</td>
					</tr>
					<tr>
						<td class="m_l_sel_sep" colspan="2"></td>
					</tr>
					<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('ent');">
						<td class="m_l_list_p m_l_list_t"><b>3</b></td>
						<td class="m_l_list_t">Remaining Period Entitlements</td>
					</tr>
					<tr>
						<td class="m_l_sel_sep" colspan="2"></td>
					</tr>
					<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('number');">
						<td class="m_l_list_p m_l_list_t"><b>4</b></td>
						<td class="m_l_list_t">Number Crunching</td>
					</tr>
					<%
					for i = 1 to 7
						if (i = 1) or (i = 7) then
							 if var_est_enabled_weekends = "0" then
							 else
					%>
					<tr>
						<td class="m_l_sel_sep" colspan="2"></td>
					</tr>
					<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('<%=i%>');">
						<td class="m_l_list_p m_l_list_t"><b><%=i + 4%></b></td>
						<td class="m_l_list_t"><b>Detailed</b> Cover Summary For <%=getdow((DateAdd("d",i - 1,getweekstart(date()))))%>, <%=DateAdd("d",i - 1,getweekstart(date()))%></td>
					</tr>
					<%
							 end if
						else	
					%>
					<tr>
						<td class="m_l_sel_sep" colspan="2"></td>
					</tr>
					<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('<%=i%>');">
						<td class="m_l_list_p m_l_list_t"><b><%if var_est_enabled_weekends = "1" then%><%=i + 4%><%else%><%=i + 3%><%end if%></b></td>
						<td class="m_l_list_t"><b>Detailed</b> Cover Summary For <%=getdow((DateAdd("d",i - 1,getweekstart(date()))))%>, <%=DateAdd("d",i - 1,getweekstart(date()))%></td>
					</tr>
					<%
						end if
					next
					%>
				</table>
			</div>
			<div class="m_l_middletitle" style="width: 544px; _width: 550px; margin-bottom: 10px;" onmouseup="showdetail('absent');">Who Was Absent During The Week?</div>
			<div id="list_absent" style="display: none; padding: 0px 0px 10px 4px;">
				This Report Shows You Who Was Off During The Week, And At What Periods. To Find Out Who Covered For Them, Simply Flip To The Detailed Cover Summary For The
				Relevant Day, Found Further On In This Weekly Report.
				<hr size="1">
				<div style="padding-left: 5px;">
				<%
				RSTESTSQL = "SELECT * FROM ATTENDANCE WHERE (((Attendance.DAYDATE) Between #" & SQLDate(getweekstart(date())) & "# And #" & SQLDate(getweekend(date())) & "#))"
				
				Set RSTEST = Server.CreateObject("Adodb.RecordSet")
				RSTEST.Open RSTESTSQL, dataconn, adopenkeyset, adlockoptimistic
				
				if RSTEST.RECORDCOUNT = 0 then
				%>				
				<b>No Staff Are Listed As Being Absent For This Week!</b>
				<%
				else
				end if
				
				RSTEST.close
				set RSTEST = nothing

				RSUSERSQL = "SELECT ID, FN, LN, TITLE FROM TIMETABLES ORDER BY LN ASC"

				Set RSUSER = Server.CreateObject("Adodb.RecordSet")
				RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
		
				if RSUSER.RECORDCOUNT = 0 then
				%>
				There Are No Staff Members In The System's Database!
				<%
				else
					do until RSUSER.EOF
						RSATTSQL = "SELECT * FROM ATTENDANCE WHERE (((Attendance.DAYDATE) Between #" & SQLDate(getweekstart(date())) & "# And #" & SQLDate(getweekend(date())) & "#) And USER = " & RSUSER("ID") & ") ORDER BY DAYDATE ASC"

						Set RSATT = Server.CreateObject("Adodb.RecordSet")
						RSATT.Open RSATTSQL, dataconn, adopenkeyset, adlockoptimistic

						if RSATT.RECORDCOUNT = 0 then
						else
						%>
				<div style="background-color: #FAD9AB; border-bottom: 1px solid #CD5301; padding-left: 5px; line-height: 24px; font-family: Tahoma,Sans-Serif;">
					<b><%=RSUSER("TITLE")%>. <%=left(RSUSER("FN"),1)%>. <%=RSUSER("LN")%></b>
				</div>
						<%
							do until RSATT.EOF
								RSDAYSQL = "SELECT * FROM Periods WHERE DAYID = " & RSATT("DAY")

								Set RSDAY = Server.CreateObject("Adodb.RecordSet")
								RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic
				%>
				<table style="padding-bottom: 5px;" cellpadding="0" cellspacing="0">
					<tr>
						<td style="width: 20px;"></td>
						<td style="padding-top: 5px; padding-bottom: 5px; font-family: Tahoma,Sans-Serif; font-size: 10pt; font-weight: bold;" colspan="<%=RSDAY("Totals")%>"><%=GetDOW(RSATT("DAY"))%>, <%=RSATT("DAYDATE")%></td>
					</tr>
					<tr>
						<td style="width: 20px;"></td>
						<td>
							<table cellpadding="0" cellspacing="0">
								<tr>
									<%
									for k = 1 to RSDAY("TOTALS")
									%>
									<td style="height: 25px;<%if k = 1 then%>border-left: 1px solid #D99251; border-right: 1px solid #D99251; border-top: 1px solid #D99251; border-bottom: 1px solid #D99251;<%elseif k = RSDAY("TOTALS") then%> border-left: 1px soild #D99251; border-right: 1px solid #D99251; border-top: 1px solid #D99251; border-bottom: 1px solid #D99251;<%else%> border-right: 1px solid #D99251; border-top: 1px solid #D99251; border-bottom: 1px solid #D99251;<%end if%>; width: 60px; padding-left: 5px; text-align: center; font-family: Tahoma,Sans-Serif; font-size: 8pt;<%if RSATT(k & "_" & RSATT("DAY")) <> "" then%> background: url('/pt/media/tt/r.png') repeat-x;<%else%> background: url('/pt/media/tt/g.png') repeat-x;<%end if%>"><b><%=k%></b></td>
									<%
									next
									k = null
									%>
								</tr>
							</table>
						</td>
					</tr>
				</table>
				<%
								RSDAY.close
								set RSDAY = nothing
							RSATT.MOVENEXT
							loop

							RSATT.close
							set RSATT = nothing
						end if
					RSUSER.MOVENEXT
					loop
		
				end if

				RSUSER.close
				set RSUSER = nothing
				%>
				</div>
			</div>
			<div class="m_l_middletitle" style="width: 544px; _width: 550px; margin-bottom: 10px;" onmouseup="showdetail('dept');">Departmental Summary</div>
			<div id="list_dept" style="display: none; padding: 0px 0px 10px 0px;">
			<%
			RSDEPTSQL = "SELECT * FROM DEPARTMENTS ORDER BY SHORT"

			Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
			RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic
			%>
				Find Out Which Department Has Required The Most Cover Over The Course Of The Week.
				Each Department Is Listed Below, With Each's Days Required Cover And The Week Total.
				<hr size="1">
				<div style="text-align: center; font-weight: bold;">Listed Departments : <i><%=RSDEPT.RECORDCOUNT%></i></div>
				<br>
					<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
						<tr style="background-color: #F7C077;">
							<td class="m_l_list_t rep_week_t" style="padding-left: 5px; text-align: left;"><b>Department</b></td>
							<%
							for k = 1 to 7
								RSDAYSQL = "SELECT * FROM Periods WHERE DAYID = " & k

								Set RSDAY = Server.CreateObject("Adodb.RecordSet")
								RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic

								if (k = 1) or (k = 7) then
									 if var_est_enabled_weekends = "0" then
									 else
							%>
							<td class="m_l_list_t rep_week_t" style="width: 30px; text-align: center;"><b><%=left(RSDAY("DAYNAME"),1)%></b></td>
							<%
									 end if
								else	
							%>
							<td class="m_l_list_t rep_week_t" style="width: 30px; text-align: center;"><b><%=left(RSDAY("DAYNAME"),1)%></b></td>
							<%
								end if
								RSDAY.close
								set RSDAY = nothing
							next
							%>
							<td class="m_l_list_t rep_week_t" style="width: 30px; text-align: center;"><b>Tot.</b></td>
						</tr>
						<%
						do until RSDEPT.EOF
						%>
						<tr>
							<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
						</tr>
						<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);">
							<td class="m_l_list_t rep_week_t" style="padding-left: 5px; text-align: left;"><%=RSDEPT("FULL")%></td>
							<%
							k = null
							for k = 1 to 7
					
								if (k = 1) or (k = 7) then
									 if var_est_enabled_weekends = "0" then
									 else
								RSCOUNTSQL = "SELECT * FROM Timetables WHERE DEPT = " & RSDEPT("DEPTID")
						
								Set RSCOUNT = Server.CreateObject("Adodb.RecordSet")
								RSCOUNT.Open RSCOUNTSQL, dataconn, adopenkeyset, adlockoptimistic
							
								do until RSCOUNT.EOF
									RSCOUNT2SQL = "SELECT * FROM Cover WHERE FOR = " & RSCOUNT("ID") & " AND DAY = " & k & " AND DAYDATE = #" & SQLDate(DateAdd("d",k - 1,getweekstart(date()))) & "#"
					
									Set RSCOUNT2 = Server.CreateObject("Adodb.RecordSet")
									RSCOUNT2.Open RSCOUNT2SQL, dataconn, adopenkeyset, adlockoptimistic
									daytot = daytot + RSCOUNT2.RECORDCOUNT
									depttotal = depttotal + RSCOUNT2.RECORDCOUNT
									RSCOUNT.MOVENEXT
									loop
							%>
							<td class="m_l_list_t rep_week_t" style="width: 30px; text-align: center;"><%=daytot%></td>
							<%
									 end if
								daytot = 0
								else

								RSCOUNTSQL = "SELECT * FROM Timetables WHERE DEPT = " & RSDEPT("DEPTID")
					
								Set RSCOUNT = Server.CreateObject("Adodb.RecordSet")
								RSCOUNT.Open RSCOUNTSQL, dataconn, adopenkeyset, adlockoptimistic
							
								do until RSCOUNT.EOF
									RSCOUNT2SQL = "SELECT * FROM Cover WHERE FOR = " & RSCOUNT("ID") & " AND DAY = " & k & " AND DAYDATE = #" & SQLDate(DateAdd("d",k - 1,getweekstart(date()))) & "#"
						
									Set RSCOUNT2 = Server.CreateObject("Adodb.RecordSet")
									RSCOUNT2.Open RSCOUNT2SQL, dataconn, adopenkeyset, adlockoptimistic
									daytot = daytot + RSCOUNT2.RECORDCOUNT
									depttotal = depttotal + RSCOUNT2.RECORDCOUNT
									RSCOUNT.MOVENEXT
									loop
							%>
							<td class="m_l_list_t rep_week_t" style="width: 30px; text-align: center;"><%=daytot%></td>
							<%
								daytot = 0
								end if
							next
							%>
							<td class="m_l_list_t rep_week_t" style="width: 30px; text-align: center;"><b><%=depttotal%></b></td>
						</tr>
						<tr>
							<td class="m_l_sel_sep" style="height: 3px;" colspan="3"></td>
						</tr>
						<%
						daytot = 0
						depttotal = 0
						RSDEPT.MOVENEXT
						loop
						%>
						<tr>
							<td class="m_l_sel_sep" colspan="2"></td>
						</tr>
					</table>
			<%
			RSDEPT.close
			set RSDEPT = nothing
			%>
			</div>
			<div class="m_l_middletitle" style="width: 544px; _width: 550px; margin-bottom: 10px;" onmouseup="showdetail('ent');">Remaining Period Entitlements</div>
			<div id="list_ent" style="display: none; padding: 0px 0px 10px 0px;">
			<%
			RSUSERSQL = "SELECT ID, FN, LN, TITLE, ENTITLEMENT, DEPT FROM Timetables ORDER BY LN ASC;"

			Set RSUSER = Server.CreateObject("Adodb.RecordSet")
			RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
			%>
				Which Poor Soul Has Had To Cover Lots Of Classes? Find Out Here!<br>
				By Looking At This Report, You Will Be Able To Find Out Who Has Been Overloaded With PleaseTakes.
				You'll Then Be Able To Make Choices As To Who Will Cover Classes For The Future.<br>
				<b>This Report DOES NOT Include Outside Cover Staff.</b><br>
				Staff Who Have Had Their Entitlement Used Up Are Listed In <b>Bold</b>.
				<hr size="1">
				<div style="text-align: center; font-weight: bold;">Listed Staff Members : <i><%=RSUSER.RECORDCOUNT%></i></div>
				<br>
				<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
					<tr style="background-color: #F7C077;">
						<td class="m_l_list_t rep_week_t" style="width: 70%; padding-left: 5px; text-align: left;"><b>Staff Member</b></td>
						<td class="m_l_list_t rep_week_t" style="width: 15%; text-align: center;"><b>Total Ent.</b></td>
						<td class="m_l_list_t rep_week_t" style="width: 15%; text-align: center;"><b>Free Ent.</b></td>
					</tr>
					<tr>
						<td class="m_l_sel_sep" style="height: 5px;" colspan="3"></td>
					</tr>
					<%
					do until RSUSER.EOF

						RSDEPTSQL = "SELECT * FROM DEPARTMENTS WHERE DEPTID = " & RSUSER("DEPT")

						Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
						RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic

						RSCOVERSQL = "SELECT * FROM COVER WHERE (((Cover.DAYDATE) Between #" & SQLDate(getweekstart(date())) & "# And #" & SQLDate(getweekend(date())) & "#) And COVERING = " & RSUSER("ID") & ")"

						Set RSCOVER = Server.CreateObject("Adodb.RecordSet")
						RSCOVER.Open RSCOVERSQL, dataconn, adopenkeyset, adlockoptimistic
					%>
					<tr <%if (RSUSER("ENTITLEMENT") - RSCOVER.RECORDCOUNT = 0) and (RSUSER("ENTITLEMENT") <> 0) then%>onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);"<%else%>onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);"<%end if%>>
						<td class="m_l_list_t rep_week_t" style="width: 70%; text-align: left; padding-left: 5px;<%if (RSUSER("ENTITLEMENT") - RSCOVER.RECORDCOUNT = 0) and (RSUSER("ENTITLEMENT") <> 0) then%> font-weight: bold;<%else%><%end if%>"><%=RSUSER("TITLE")%>. <%=left(RSUSER("FN"),1)%>. <%=RSUSER("LN")%>, <span style="font-size: 10pt;"><%=RSDEPT("FULL")%></span></td>
						<td class="m_l_list_t rep_week_t" style="width: 15%; text-align: center;<%if RSUSER("ENTITLEMENT") - RSCOVER.RECORDCOUNT = 0 and (RSUSER("ENTITLEMENT") <> 0) then%> font-weight: bold;<%else%><%end if%>"><%=RSUSER("ENTITLEMENT")%></td>
						<td class="m_l_list_t rep_week_t" style="width: 15%; text-align: center;<%if RSUSER("ENTITLEMENT") - RSCOVER.RECORDCOUNT = 0 and (RSUSER("ENTITLEMENT") <> 0) then%> font-weight: bold;<%else%><%end if%>"><%=RSUSER("ENTITLEMENT") - RSCOVER.RECORDCOUNT%></td>
					</tr>
					<tr>
						<td class="m_l_sel_sep" style="height: 5px;" colspan="3"></td>
					</tr>
					<%
						RSDEPT.close
						RSCOVER.close
						set RSDEPT = nothing
						set RSCOVER = nothing
					RSUSER.MOVENEXT
					loop
					%>
					<tr>
						<td class="m_l_sel_sep" style="height: 5px;" colspan="3"></td>
					</tr>
				</table>
				<%
				RSUSER.close
				set RSUSER = nothing
				%>
			</div>
			<div class="m_l_middletitle" style="width: 544px; _width: 550px; margin-bottom: 10px;" onmouseup="showdetail('number');">Number Crunching</div>
			<div id="list_number" style="display: none; padding: 0px 0px 10px 4px;">
			<%
			RSCOVERSQL = "SELECT ID FROM COVER WHERE (((Cover.DAYDATE) Between #" & SQLDate(getweekstart(date())) & "# And #" & SQLDate(getweekend(date())) & "#))"

			Set RSCOVER = Server.CreateObject("Adodb.RecordSet")
			RSCOVER.Open RSCOVERSQL, dataconn, adopenkeyset, adlockoptimistic

			RSOCOVERSQL = "SELECT ID FROM COVER WHERE (((Cover.DAYDATE) Between #" & SQLDate(getweekstart(date())) & "# And #" & SQLDate(getweekend(date())) & "#) And OCOVER <> 0)"

			Set RSOCOVER = Server.CreateObject("Adodb.RecordSet")
			RSOCOVER.Open RSOCOVERSQL, dataconn, adopenkeyset, adlockoptimistic

			RSABSENTSQL = "SELECT USER FROM ATTENDANCE WHERE (((Attendance.DAYDATE) Between #" & SQLDate(getweekstart(date())) & "# And #" & SQLDate(getweekend(date())) & "#)) GROUP BY USER"

			Set RSABSENT = Server.CreateObject("Adodb.RecordSet")
			RSABSENT.Open RSABSENTSQL, dataconn, adopenkeyset, adlockoptimistic
			%>
				This Report Gives You One Thing: Figures. In A Plain And Simple Way, This Report Explains To You Exactly What Happened Over The Week, With Only The Simple Numbers That Are Required.
				<hr size="1">
				<div style="padding-left: 5px; text-align: center;">
					<span style="font-size: 16pt;">The Total Number Of PleaseTakes Over The Week:</span><br>
					<span style="font-size: 72pt; letter-spacing: -8px;"><%=RSCOVER.RECORDCOUNT%></span>
					<hr size="1">
					<span style="font-size: 16pt;">The Total Number Of Absent Staff:</span><br>
					<span style="font-size: 72pt; letter-spacing: -8px;"><%=RSABSENT.RECORDCOUNT%></span>
					<hr size="1">
					<span style="font-size: 16pt;">The Number Of Outside Cover PleaseTakes:</span><br>
					<span style="font-size: 72pt; letter-spacing: -8px;"><%=RSOCOVER.RECORDCOUNT%></span>
				</div>
			<%
			RSCOVER.close
			RSOCOVER.close
			RSABSENT.close
			set RSCOVER = nothing
			set RSOCOVER = nothing
			set RSABSENT = nothing
			%>
			</div>
			<%
			for j = 1 to 7
				if (j = 1) or (j = 7) then
					 if var_est_enabled_weekends = "0" then
					 else
			%>
			<div class="m_l_middletitle" style="width: 544px; _width: 550px; margin-bottom: 10px;" onmouseup="showdetail('<%=j%>');">Detailed Cover Summary For <%=getdow((DateAdd("d",j - 1,getweekstart(date()))))%>, <%=DateAdd("d",j - 1,getweekstart(date()))%></div>
					<div id="list_<%=j%>" style="display: none; padding: 0px 0px 0px 4px;">
					<!--#include virtual="/pt/modules/ss/timetables/admin_rep_wk_sum.inc"-->
					<div style="height: 10px;"></div>
			</div>
			<%
					 end if
				else	
			%>
			<div class="m_l_middletitle" style="width: 544px; _width: 550px; margin-bottom: 10px;" onmouseup="showdetail('<%=j%>');">Detailed Cover Summary For <%=getdow((DateAdd("d",j - 1,getweekstart(date()))))%>, <%=DateAdd("d",j - 1,getweekstart(date()))%></div>
					<div id="list_<%=j%>" style="display: none; padding: 0px 0px 0px 4px;">
					<!--#include virtual="/pt/modules/ss/timetables/admin_rep_wk_sum.inc"-->
					<div style="height: 10px;"></div>
			</div>
			<%
				end if
			next
			%>
		</div>

		<div style="line-height: 10px;">&nbsp;</div>
	</div>
<%
	end if
elseif pagetype = "5" then
	if (request("type") = "2") then
		if (request("print") = "1") then
			if (request("printout") = "") then
				response.redirect "/pt/admin/print.asp?id=7"
			else
				response.redirect "/pt/admin/print.asp?id=7&uid=" & request("printout")
			end if
		else
%>
	<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->
	<div class="m_l">
		<div class="m_l_title">View/Print Staff Timetables</div>
		<div class="m_l_subtitle">View And Print Staff Timetables.</div>
		<div class="m_l_ins">To Print ALL Timetables, Please Click <a href="#" onmouseup="location.href='reports.asp?id=5&amp;type=2&amp;print=1'"><img src="/pt/media/icons/16_print.gif" border="0" alt="Click Here To Print Staff Timetables.">&nbsp;Here.</a></div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<%
			RSCHECKSQL = "SELECT ID, FN, LN, TITLE FROM TIMETABLES ORDER BY LN ASC"

			Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
			RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
			
			if RSCHECK.RECORDCOUNT = 0 then
				response.redirect "/pt/admin/reports.asp?id=5"
			else
			%>
			<div style="padding-bottom: 10px;">
				All Your Staff Timetables Are Listed Below, In Alphabetical Order. To View One, Just Click The Relevant Name. If You Wish To Print ALL Timetables, Please Click <a href="#" onmouseup="location.href='reports.asp?id=5&amp;type=2&amp;print=1'"><b>Here</b></a> Or "Print" Above.
				To Print Individual Timetables, Click The Relevant Member Of Staff, And Then "Print Timetable".
				<br>
				<b>Please Be Patient While All The Timetables Are Loaded.</b>
				<hr size="1">
				<div style="width: 100%; text-align: center; font-weight: bold;"><i>Total Number Of Timetables: <%=RSCHECK.RECORDCOUNT%></i></div>
			</div>
			<%
				do until RSCHECK.EOF
			%>
			<div class="m_l_middletitle" style="width: 546px; _width: 100%; margin-bottom: 10px; font-size: 12pt;" onmouseup="showdetail('tt_<%=RSCHECK("ID")%>')"><%=RSCHECK("LN")%>, <%=left(RSCHECK("FN"),1)%>.</div>
			<div id="list_tt_<%=RSCHECK("ID")%>" style="width: 540px; _width: 100%; padding: 5px 0px 10px 0px; display: none; margin-top: -10px; line-height: 28px;">
			<%
			if var_est_enabled_weekends = "1" then
			%>
			<!--#include virtual="/pt/modules/ss/timetables/admin_rep_staff_7day.inc"-->
			<%
			else
			%>
			<!--#include virtual="/pt/modules/ss/timetables/admin_rep_staff_5day.inc"-->
			<%
			end if
			%>
				<div style="font-size: 10pt; font-weight: bold;">
					<a href="#" onmouseup="location.href='reports.asp?id=5&amp;type=2&amp;print=1&amp;printout=<%=RSCHECK("ID")%>'">Click Here To Print <%=left(RSCHECK("FN"),1)%>. <%=RSCHECK("LN")%>'s Timetable</a>
				</div>
			</div>
			<%
				RSCHECK.MOVENEXT
				loop
			end if

			RSCHECK.close
			set RSCHECK = nothing
			%>
		</div>
	</div>
<%
		end if
	elseif (request("type") = "3") then

	RSUSERSQL = "SELECT ID, LN, FN, TITLE, DEPT, CATEGORY, ENTITLEMENT, DEFROOM FROM Timetables ORDER BY ENTITLEMENT DESC, LN"

	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<div class="m_l">
		<div class="m_l_title">Period Entitlements</div>
		<div class="m_l_subtitle">How Many Periods Staff Get A Week For Please Takes.</div>
		<div class="m_l_ins">For A Printer-Friendly Version Of This List, Please Click <a href="staff.asp?id=4&amp;print=1"><img border="0" alt="Click Here To Print The Entitlement List Below." src="/pt/media/icons/16_print.gif"> Here</a>.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			If Any Details Are Incorrect, Please Go <a href="staff.asp?id=4">Here</a> To Amend Them.
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep" colspan="3"></td>
				</tr>
				<%
				do until RSUSER.EOF
					RSDEPTSQL = "SELECT * FROM DEPARTMENTS WHERE DEPTID = " & RSUSER("DEPT")

					Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
					RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic
				%>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="<%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%>"></td>
					<td class="m_l_list_t" style="width: 80px;"><b><%=RSUSER("ENTITLEMENT")%></b> <%if RSUSER("ENTITLEMENT") = "1" then%>Period<%else%>Periods<%end if%></td>
					<td class="m_l_list_t"><b><%=RSUSER("LN")%>, <%=Left(RSUSER("FN"),1)%>.</b> :: 
					<%
					if RSUSER("CATEGORY") = "T" then
					response.write "Teacher"
					elseif RSUSER("CATEGORY") = "PT" then
					response.write "Principal Teacher"
					elseif RSUSER("CATEGORY") = "DHT" then
					response.write "Deputy Head Teacher"
					elseif RSUSER("CATEGORY") = "HT" then
					response.write "Head Teacher"
					elseif RSUSER("CATEGORY") = "PS" then
					response.write "Pupil Support"
					elseif RSUSER("CATEGORY") = "OC" then
					response.write "Outside Cover"
					else
					response.write "Unknown Position"
					end if
					%>, <%=RSDEPT("FULL")%>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="3"></td>
				</tr>
				<%
					RSDEPT.close
					set RSDEPT = nothing
				RSUSER.MOVENEXT
				loop
				%>
			</table>
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 10pt;">Displaying A Total Of <b><%=RSUSER.RECORDCOUNT%></b> Staff Records</div>
		</div>
	</div>
<%
				RSUSER.close
				set RSUSER = nothing
	else
%>
	<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->
	<div class="m_l">
		<div class="m_l_title">View Staff Information</div>
		<div class="m_l_subtitle">View And Print Everything You Need To Know About Your Staff.</div>
		<div class="m_l_ins">Please Choose A Report From Below.</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<%
			RSCHECKSQL = "SELECT ID FROM TIMETABLES"

			Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
			RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
			
			if RSCHECK.RECORDCOUNT = 0 then
			%>
			There Are No Staff Listed In The System's Database, So No Reports Can Be Generated!
			<div class="botopts">
				<ul>
					<li><a href="staff.asp?id=6">Add A Member Of Staff</a></li>
					<li><a href="default.asp?id=1">Return To The Admin Homepage</a></li>
				</ul>
			</div>			
			<%
			else
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_t_l"><a href="reports.asp?id=5&amp;type=2"><img src="/pt/media/icons/48_cal.png" border="0" alt="Click Here To View Or Print Staff Timetables."></a></td>
					<td class="m_l_sel_t_r"><a href="reports.asp?id=5&amp;type=2">View/Print Staff Timetables</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="reports.asp?id=5&amp;type=3"><img src="/pt/media/icons/48_slips.png" border="0" alt="Click Here To View Or Print The Staff Entitlements List."></a></td>
					<td class="m_l_sel_t_r"><a href="reports.asp?id=5&amp;type=3">View/Print Period Entitlements List</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="reports.asp?id=5&amp;type=4"><img src="/pt/media/icons/48_shares.png" border="0" alt="Click Here To View Any Job Shares."></a></td>
					<td class="m_l_sel_t_r"><a href="reports.asp?id=5&amp;type=4">View Job Shares</a></td>
				</tr>
			</table>
			<%
			end if
			
			RSCHECK.close
			set RSCHECK = nothing
			%>
		</div>
	</div>
<%
	end if
else
%>
	<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->
	<div class="m_l">
		<div class="m_l_title">View Reports</div>
		<div class="m_l_subtitle">View And Print Everything You Need To Know.</div>
		<div class="m_l_ins">Please Choose A Type Of Report From Below...</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_t_l"><a href="reports.asp?id=2"><img src="/pt/media/icons/48_slips.png" border="0" alt="View Today's Slips For Staff Notifying Them Of A Class To Cover. These Can Also Be Printed."></a></td>
					<td class="m_l_sel_t_r" style="font-size: 14pt;"><a href="reports.asp?id=2">View/Print PleaseTake Slips</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="reports.asp?id=3"><img src="/pt/media/icons/48_view.png" border="0" alt="View All Of The PleaseTakes In One Table, For The Day Of Your Choice."></a></td>
					<td class="m_l_sel_t_r" style="font-size: 14pt;"><a href="reports.asp?id=3">View The Cover Summary</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="reports.asp?id=4"><img src="/pt/media/icons/48_cal.png" border="0" alt="View The Weekly Report, Showing You Everything That Has Happened Over The Past Week."></a></td>
					<td class="m_l_sel_t_r"><a href="reports.asp?id=4">View Weekly Report</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>				
				<tr>
					<td class="m_l_sel_t_l"><a href="reports.asp?id=5"><img src="/pt/media/icons/48_book.png" border="0" alt="Click To See Everyone's Timetables And More."></a></td>
					<td class="m_l_sel_t_r"><a href="reports.asp?id=5">View Staff Information</a></td>
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