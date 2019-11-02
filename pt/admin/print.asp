<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" >
<!--
Print ID's
	(2) = Printing Today's PleaseTake Slips
	(3) = Printing Cover Summary
	(4) = Printing Weekly Report
	(5) = Printing Weekly Report (From Backup)
	(6) = Printing Entitlement List
	(7) = Printing Staff Timetables (Reports)
	(8) = Printing Individual PleaseTake Slip
-->
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
<script language="javascript" type="text/javascript" src="/pt/modules/js/admin.js"></script>
<title><%=var_ptitle%></title>
</head>

<body onload="print(); history.back();">

<div class="main" style="width: 648px; color: #000;">
<%
if pagetype = "2" then
	daydow = request("dow")
	daydate = request("daydate")
%>
	<div style="font-size: 18pt; font-weight: bold; letter-spacing: -2px;">PleaseTakes <%=var_ver%></div>
	<hr size="1">
	<div style="font-size: 10pt; padding-left: 10pt; font-weight: bold;">
		PleaseTake Slips For <%=getdow(daydow)%>, <%=day(daydate)%>&nbsp;<%=cal_getmonth(month(daydate))%>&nbsp;<%=year(daydate)%>, Printed At <%=date%>, <%=time%>
	</div>
	<hr size="1">
	<div style="font-size: 12pt; padding-left: 10pt; font-weight: bold;">
		STAFF COVER
	</div>
	<hr size="1">
			<%
			RSSLIPSSQL = "SELECT * FROM Cover WHERE DAY = " & daydow & " AND DAYDATE = #" & SQLDate(daydate) & "# AND OCOVER <> 1 ORDER BY PERIOD"
		
			Set RSSLIPS = Server.CreateObject("Adodb.RecordSet")
			RSSLIPS.Open RSSLIPSSQL, dataconn, adopenkeyset, adlockoptimistic
			
			if RSSLIPS.RECORDCOUNT = 0 then
				RSSLIPS.close
				set RSSLIPS = nothing
			%>
			There Are No Staff Cover Slips To Be Printed For <%=getdow(daydow)%>, <%=daydate%>!
			<%
			else

			nopages = round((RSSLIPS.RECORDCOUNT / 4))
			leftover = RSSLIPS.RECORDCOUNT mod 4
			
			if leftover <> 0 then
				nopages = nopages + 1
			else
				nopages = nopages
			end if
	
			for p = 1 to nopages
				if p = nopages then
			%>
			<div>
			<%
				else
			%>
			<div style="page-break-after: always;">
			<%
				end if

				for s = 1 to 4
			
				if RSSLIPS.EOF then
				else

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
			<div class="rep_slip_p">
				<div class="rep_slip_c_p">
					<div style="height: 22px; line-height: 15px; border-bottom: 1px solid #000;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr>
								<td style="width: 50%; color: #000; font-size: 10pt; font-family: Tahoma,Sans-Serif;"><%=getdow(daydow)%>, <%=day(daydate)%>&nbsp;<%=cal_getmonth(month(daydate))%>&nbsp;<%=year(daydate)%></td>
								<td style="width: 50%; color: #000; font-size: 10pt; text-align: right; font-family: Tahoma, Sans-Serif;"><b><%=var_pname%></b> <%=var_ver%></td>
							</tr>
						</table>
					</div>
					<div class="rep_slip_b_p">
					<table class="m_l_sel_t" style="height: 100%;" cellpadding="0" cellspacing="0">
						<tr>
							<td style="width: 60%;">
							<table class="m_l_sel_t" style="height: 100%;" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="2" class="rep_slip_t_p" style="border-bottom: 1px solid #000;"><b>Please Arrange To Take This Class.</b></td>
								</tr>
								<tr>
									<td class="rep_slip_t_p" style="height: 30px; width: 25%; border-bottom: 1px solid #000;">&nbsp;<b>Covering:&nbsp;&nbsp;</b></td>
									<td class="rep_slip_t_p" style="border-bottom: 1px solid #000;"><b><%=left(RSCOVERING("FN"),1)%>.&nbsp;<%=RSCOVERING("LN")%></b> (<%=RSCOVERINGDEPT("SHORT")%><%if RSCOVERING("DEFROOM") <> "" then%>, Room <b><%=RSCOVERING("DEFROOM")%></b><%else%><%end if%>)</td>
								</tr>
								<tr>
									<td class="rep_slip_t_p" style="height: 30px; width: 25%; border-bottom: 1px solid #000;">&nbsp;<b>For:</b></td>
									<td class="rep_slip_t_p" style="border-bottom: 1px solid #000;"><%=left(RSFOR("FN"),1)%>.&nbsp;<%=RSFOR("LN")%> (<%=RSFORDEPT("SHORT")%>)</td>
								</tr>
								<tr>
									<td class="rep_slip_t_p" style="height: 30px; width: 25%; border-bottom: 1px solid #000;">&nbsp;<b>Class:</b></td>
									<td class="rep_slip_t_p" style="border-bottom: 1px solid #000;"><%=RSCLASS(RSSLIPS("PERIOD") & "_" & daydow)%></td>
								</tr>
								<tr>
									<td colspan="2" class="rep_slip_t_p" style="text-align: right;"><b>Thank You Very Much!</b>&nbsp;</td>
								</tr>
							</table>
							</td>
							<td style="width: 40%; text-align: right;">
							<table class="m_l_sel_t" style="height: 100%;" cellpadding="0" cellspacing="0">
								<tr>
									<td class="rep_slip_t_p" style="width: 50%; height: 27px; border-bottom: 1px solid #000; border-left: 1px solid #000; border-right: 1px solid #000; text-align: center;"><b>Period</b></td>
									<td class="rep_slip_t_p" style="width: 50%; height: 27px; border-bottom: 1px solid #000; text-align: center;"><b>Room</b></td>
								</tr>
								<tr>
									<td class="rep_slip_t_p" style="border-left: 1px solid #000; border-right: 1px solid #000; text-align: center; font-size: 70pt;"><%=RSSLIPS("PERIOD")%></td>
									<td class="rep_slip_t_p" style="text-align: center; font-size: 36pt; letter-spacing: -6px;"><%=RSCLASS("R" & RSSLIPS("PERIOD") & "_" & daydow)%></td>
								</tr>
							</table>
							</td>
						</tr>
					</table>
					</div>
				</div>
			</div>
			<hr size="1">
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
				end if

				next
				%>
			</div>
			<%
			next
			RSSLIPS.close
			set RSSLIPS = nothing
			end if

			%>
	<hr size="1" style="page-break-before: always;">
	<div style="font-size: 12pt; padding-left: 10pt; font-weight: bold;">
		OUTSIDE COVER
	</div>
	<hr size="1">
			<%
			RSSLIPSSQL = "SELECT * FROM Cover INNER JOIN OCover ON Cover.COVERING = OCover.ID WHERE (((Cover.DAY)=" & daydow & ") AND ((Cover.DAYDATE)=#" & SQLDate(daydate) & "#) AND ((Cover.OCOVER)=1)) ORDER BY OCover.LN, Cover.Period"
		
			Set RSSLIPS = Server.CreateObject("Adodb.RecordSet")
			RSSLIPS.Open RSSLIPSSQL, dataconn, adopenkeyset, adlockoptimistic
			
			if RSSLIPS.RECORDCOUNT = 0 then
				RSSLIPS.close
				set RSSLIPS = nothing
			%>
			There Are No Outside Cover Slips To Be Printed For <%=getdow(daydow)%>, <%=daydate%>!
			<%
			else

			nopages = round((RSSLIPS.RECORDCOUNT / 4))
			leftover = RSSLIPS.RECORDCOUNT mod 4
			
			if leftover <> 0 then
				nopages = nopages + 1
			else
				nopages = nopages
			end if
	
			for p = 1 to nopages
				if p = nopages then
			%>
			<div>
			<%
				else
			%>
			<div style="page-break-after: always;">
			<%
				end if

				for s = 1 to 4
				
				if RSSLIPS.EOF then
				else

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
		
				RSCOVERINGSQL = "SELECT * FROM OCover WHERE ID = " & RSSLIPS("COVERING")
				Set RSCOVERING = Server.CreateObject("Adodb.RecordSet")
				RSCOVERING.Open RSCOVERINGSQL, dataconn, adopenkeyset, adlockoptimistic
			%>
			<div class="rep_slip_p">
				<div class="rep_slip_c_p">
					<div style="height: 22px; line-height: 15px; border-bottom: 1px solid #000;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr>
								<td style="width: 50%; color: #000; font-size: 10pt; font-family: Tahoma,Sans-Serif;"><%=getdow(daydow)%>, <%=day(daydate)%>&nbsp;<%=cal_getmonth(month(daydate))%>&nbsp;<%=year(daydate)%></td>
								<td style="width: 50%; color: #000; font-size: 10pt; text-align: right; font-family: Tahoma, Sans-Serif;"><b><%=var_pname%></b> <%=var_ver%></td>
							</tr>
						</table>
					</div>
					<div class="rep_slip_b_p">
					<table class="m_l_sel_t" style="height: 100%;" cellpadding="0" cellspacing="0">
						<tr>
							<td style="width: 60%;">
							<table class="m_l_sel_t" style="height: 100%;" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="2" class="rep_slip_t_p" style="border-bottom: 1px solid #000;"><b>Please Arrange To Take This Class.</b></td>
								</tr>
								<tr>
									<td class="rep_slip_t_p" style="height: 30px; width: 25%; border-bottom: 1px solid #000;">&nbsp;<b>Covering:&nbsp;&nbsp;</b></td>
									<td class="rep_slip_t_p" style="border-bottom: 1px solid #000;"><b><%=left(RSCOVERING("FN"),1)%>.&nbsp;<%=RSCOVERING("LN")%></b> (Outside Cover)</td>
								</tr>
								<tr>
									<td class="rep_slip_t_p" style="height: 30px; width: 25%; border-bottom: 1px solid #000;">&nbsp;<b>For:</b></td>
									<td class="rep_slip_t_p" style="border-bottom: 1px solid #000;"><%=left(RSFOR("FN"),1)%>.&nbsp;<%=RSFOR("LN")%> (<%=RSFORDEPT("SHORT")%>)</td>
								</tr>
								<tr>
									<td class="rep_slip_t_p" style="height: 30px; width: 25%; border-bottom: 1px solid #000;">&nbsp;<b>Class:</b></td>
									<td class="rep_slip_t_p" style="border-bottom: 1px solid #000;"><%=RSCLASS(RSSLIPS("PERIOD") & "_" & daydow)%></td>
								</tr>
								<tr>
									<td colspan="2" class="rep_slip_t_p" style="text-align: right;"><b>Thank You Very Much!</b>&nbsp;</td>
								</tr>
							</table>
							</td>
							<td style="width: 40%; text-align: right;">
							<table class="m_l_sel_t" style="height: 100%;" cellpadding="0" cellspacing="0">
								<tr>
									<td class="rep_slip_t_p" style="width: 50%; height: 27px; border-bottom: 1px solid #000; border-left: 1px solid #000; border-right: 1px solid #000; text-align: center;"><b>Period</b></td>
									<td class="rep_slip_t_p" style="width: 50%; height: 27px; border-bottom: 1px solid #000; text-align: center;"><b>Room</b></td>
								</tr>
								<tr>
									<td class="rep_slip_t_p" style="border-left: 1px solid #000; border-right: 1px solid #000; text-align: center; font-size: 70pt;"><%=RSSLIPS("PERIOD")%></td>
									<td class="rep_slip_t_p" style="text-align: center; font-size: 36pt; letter-spacing: -6px;"><%=RSCLASS("R" & RSSLIPS("PERIOD") & "_" & daydow)%></td>
								</tr>
							</table>
							</td>
						</tr>
					</table>
					</div>
				</div>
			</div>
			<hr size="1">
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

				RSSLIPS.MOVENEXT
				end if

				next
				%>
			</div>
			<%
			next
			end if
			
			RSSLIPS.close
			set RSSLIPS = nothing

elseif pagetype = "3" then
	daydow = request("dow")
	daydate = request("daydate")
%>
	<div style="font-size: 22pt; font-weight: bold; letter-spacing: -2px;">PleaseTakes <%=var_ver%></div>
	<hr size="1">
	<div style="font-size: 12pt; font-weight: bold; padding-left: 10pt;">
		Cover Summary For <%=getdow(daydate)%>, <%=daydate%><br>
		Printed On <%=display_date%> At <%=display_time%>
	</div>
	<hr size="1">
	<div style="font-size: 10pt;">
		<b>Legend:</b>
		<div style="padding-left: 30px;">
			<b>NC</b> - No Class<br>
			<b>P</b> - Is Present, And Able To Take Own Class<br>
			<b>?</b> - The Staff Member Has Been Selected As Absent, But No Cover Is Arranged!<br>
			<b>Other Staff Name</b> - Staff Member Is Absent, So This Staff Member Is Covering The Class
		</div>
	</div>
	<p>
	<!--#include virtual="/pt/modules/ss/timetables/admin_rep_sum_p.inc"-->
<%
elseif pagetype = "4" then
%>
<table style="height: 1028px; width: 100%;" cellpadding="0" cellspacing="0">
	<tr>
		<td style="vertical-align: middle; text-align: right;">
			<img src="/pt/media/login/<%=var_est_logimg%>" border="0" alt="<%=var_est_full%>"><br>
			<img src="/pt/media/admin/p_wr.gif" border="0" alt="The Weekly Report"><br>
			<span style="color: #000; font-size: 18pt; font-weight: bold; letter-spacing: -2px;">
				For <%=var_est_full%><br>
				Week <%=GetWeekStart(date())%> - <%=GetWeekEnd(date())%><br>
				Using PleaseTakes <%=var_ver%><br>
			</span>
		</td>
	</tr>
</table>
<div style="page-break-after: always;">
	<div style="font-size: 22pt; font-weight: bold; letter-spacing: -2px;">PleaseTakes Report Contents</div>
	<hr size="1">
	<div style="padding-left: 15px;">
		<table class="m_l_sel_t" style="color: #000;" cellpadding="0" cellspacing="0">
			<tr>
				<td style="width: 100px; font-size: 48pt; font-weight: bold; text-align: center; letter-spacing: -10px;">1</td>
				<td>
					<table class="m_l_sel_t" style="color: #000;" cellpadding="0" cellspacing="0">
						<tr>
							<td style="height: 50%; font-weight: bold;">Who Was Absent During The Week?</td>
						</tr>
						<tr>
							<td style="height: 50%; font-size: 10pt;">See Who Was Off And When At A Glance.</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
			</tr>
			<tr>
				<td style="width: 100px; font-size: 48pt; font-weight: bold; text-align: center; letter-spacing: -10px;">2</td>
				<td>
					<table class="m_l_sel_t" style="color: #000;" cellpadding="0" cellspacing="0">
						<tr>
							<td style="height: 50%; font-weight: bold;">Departmental Statistics</td>
						</tr>
						<tr>
							<td style="height: 50%; font-size: 10pt;">See Which Departments Had The Most Absent Staff, And On Which Days.</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
			</tr>
			<tr>
				<td style="width: 100px; font-size: 48pt; font-weight: bold; text-align: center; letter-spacing: -10px;">3</td>
				<td>
					<table class="m_l_sel_t" style="color: #000;" cellpadding="0" cellspacing="0">
						<tr>
							<td style="height: 50%; font-weight: bold;">Remaining Period Entitlements</td>
						</tr>
						<tr>
							<td style="height: 50%; font-size: 10pt;">Who Covered The Most Periods During The Week? Check Out Here.</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
			</tr>
			<tr>
				<td style="width: 100px; font-size: 48pt; font-weight: bold; text-align: center; letter-spacing: -10px;">4</td>
				<td>
					<table class="m_l_sel_t" style="color: #000;" cellpadding="0" cellspacing="0">
						<tr>
							<td style="height: 50%; font-weight: bold;">Number Crunching</td>
						</tr>
						<tr>
							<td style="height: 50%; font-size: 10pt;">See Simple, Easy To Understand Figures For The Week.</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
			</tr>
			<%
			for i = 1 to 7
				if (i = 1) or (i = 7) then
					 if var_est_enabled_weekends = "0" then
					 else
			%>
			<tr>
				<td style="width: 100px; font-size: 48pt; font-weight: bold; text-align: center; letter-spacing: -10px;"><%=i + 4%></td>
				<td>
					<table class="m_l_sel_t" style="color: #000;" cellpadding="0" cellspacing="0">
						<tr>
							<td style="height: 50%; font-weight: bold;">Detailed Cover Summary For <%=getdow((DateAdd("d",i - 1,getweekstart(date()))))%>, <%=DateAdd("d",i - 1,getweekstart(date()))%></td>
						</tr>
						<tr>
							<td style="height: 50%; font-size: 10pt;">See The Timetabled-Form Cover Summary For <%=getdow((DateAdd("d",i - 1,getweekstart(date()))))%>, <%=DateAdd("d",i - 1,getweekstart(date()))%>.</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
			</tr>
			<%
					 end if
				else	
			%>
			<tr>
				<td style="width: 100px; font-size: 48pt; font-weight: bold; text-align: center; letter-spacing: -10px;"><%if var_est_enabled_weekends = "1" then%><%=i + 4%><%else%><%=i + 3%><%end if%></td>
				<td>
					<table class="m_l_sel_t" style="color: #000;" cellpadding="0" cellspacing="0">
						<tr>
							<td style="height: 50%; font-weight: bold;">Detailed Cover Summary For <%=getdow((DateAdd("d",i - 1,getweekstart(date()))))%>, <%=DateAdd("d",i - 1,getweekstart(date()))%></td>
						</tr>
						<tr>
							<td style="height: 50%; font-size: 10pt;">See The Timetabled-Form Cover Summary For <%=getdow((DateAdd("d",i - 1,getweekstart(date()))))%>, <%=DateAdd("d",i - 1,getweekstart(date()))%>.</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
			</tr>
			<%
				end if
			next
			
			i = null
			%>
		</table>
	</div>
</div>
<div style="page-break-after: always;">
	<div style="font-size: 22pt; font-weight: bold; letter-spacing: -2px;">Who Was Absent During The Week?</div>
	<hr size="1">
	<div style="padding-left: 15px;">
		This Report Shows You Who Was Off During The Week, And At What Periods. To Find Out Who Covered For Them, Simply Flip To The Detailed Cover Summary For The
		Relevant Day, Found Further On In This Weekly Report.
		<hr size="1">
		<div style="padding-left: 5px;">
		<%
		RSUSERSQL = "SELECT ID, FN, LN, TITLE FROM TIMETABLES ORDER BY LN ASC"

		Set RSUSER = Server.CreateObject("Adodb.RecordSet")
		RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
		
		if RSUSER.RECORDCOUNT = 0 then
		%>
		There Are No Staff Members In The System's Database!
		<%
		else
			do until RSUSER.EOF
				RSATTSQL = "SELECT * FROM ATTENDANCE WHERE (((Attendance.DAYDATE) Between #" & SQLDate(getweekstart(date())) & "# And #" & SQLDate(getweekend(date())) & "#) And USER = " & RSUSER("ID") & ")"

				Set RSATT = Server.CreateObject("Adodb.RecordSet")
				RSATT.Open RSATTSQL, dataconn, adopenkeyset, adlockoptimistic

				if RSATT.RECORDCOUNT = 0 then
				else
				%>
		<div style="width: 100%;">
			<b><%=RSUSER("TITLE")%>. <%=left(RSUSER("FN"),1)%>. <%=RSUSER("LN")%></b>
		</div>
				<%
					do until RSATT.EOF
						RSDAYSQL = "SELECT * FROM Periods WHERE DAYID = " & RSATT("DAY")

						Set RSDAY = Server.CreateObject("Adodb.RecordSet")
						RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic
		%>
		<table style="color: #000;" cellpadding="0" cellspacing="0">
			<tr>
				<td style="width: 20px;"></td>
				<td style="border-bottom: 1px solid #000; font-size: 10pt;" colspan="<%=RSDAY("Totals")%>"><%=GetDOW(RSATT("DAY"))%>, <%=RSATT("DAYDATE")%></td>
			</tr>
			<tr>
				<td style="width: 20px;"></td>
				<td>
					<table style="color: #000;" cellpadding="0" cellspacing="0">
						<tr>
							<%
							for k = 1 to RSDAY("TOTALS")
							%>
							<td style="width: 60px; padding-left: 5px; <%if k = RSDAY("TOTALS") then%><%else%> border-right: 1px solid #000;<%end if%>"><%=k%> - <%if RSATT(k & "_" & RSATT("DAY")) <> "" then%>A<%else%>P<%end if%></td>
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
</div>
<div style="page-break-after: always;">
	<div style="font-size: 22pt; font-weight: bold; letter-spacing: -2px;">Departmental Statistics</div>
	<hr size="1">
	<%
	RSDEPTSQL = "SELECT * FROM DEPARTMENTS ORDER BY SHORT"

	Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
	RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic
	%>
	<div style="padding-left: 15px;">
		Find Out Which Department Has Required The Most Cover Over The Course Of The Week.
		Each Department Is Listed Below, With Each's Days Required Cover And The Week Total.
		<hr size="1">
		<div style="text-align: center; font-weight: bold;">Listed Departments : <i><%=RSDEPT.RECORDCOUNT%></i></div>
		<br>
			<table class="m_l_sel_t" style="color: #000;" cellpadding="0" cellspacing="0">
				<tr>
					<td style="border-bottom: 1px solid #000; padding-left: 5px; font-size: 16pt;"><b>Department</b></td>
					<%
					for k = 1 to 7
						RSDAYSQL = "SELECT * FROM Periods WHERE DAYID = " & k

						Set RSDAY = Server.CreateObject("Adodb.RecordSet")
						RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic

						if (k = 1) or (k = 7) then
							 if var_est_enabled_weekends = "0" then
							 else
					%>
					<td style="width: 30px; text-align: center; border-bottom: 1px solid #000; padding-left: 5px; font-size: 16pt;"><b><%=left(RSDAY("DAYNAME"),1)%></b></td>
					<%
							 end if
						else	
					%>
					<td style="width: 30px; text-align: center; border-bottom: 1px solid #000; padding-left: 5px; font-size: 16pt;"><b><%=left(RSDAY("DAYNAME"),1)%></b></td>
					<%
						end if
						RSDAY.close
						set RSDAY = nothing
					next
					%>
					<td style="width: 30px; text-align: center; border-bottom: 1px solid #000; padding-left: 5px; font-size: 16pt;"><b>Tot.</b></td>
				</tr>
				<%
				do until RSDEPT.EOF
				%>
				<tr>
					<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
				</tr>
				<tr>
					<td style="border-bottom: 1px solid #000; padding-left: 5px;"><%=RSDEPT("FULL")%></td>
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
					<td style="width: 50px; border-bottom: 1px solid #000; text-align: center;"><%=daytot%></td>
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
					<td style="width: 50px; border-bottom: 1px solid #000; text-align: center;"><%=daytot%></td>
					<%
						daytot = 0
						end if
					next
					%>
					<td style="width: 50px; border-bottom: 1px solid #000; text-align: center;"><b><%=depttotal%></b></td>
				</tr>
				<%
				daytot = 0
				depttotal = 0
				RSDEPT.MOVENEXT
				loop
				%>
			</table>
	</div>
	<%
	RSDEPT.close
	set RSDEPT = nothing
	%>
</div>
<div style="page-break-after: always;">
	<div style="font-size: 22pt; font-weight: bold; letter-spacing: -2px;">Remaining Period Entitlements</div>
	<hr size="1">
	<%
	RSUSERSQL = "SELECT ID, FN, LN, TITLE, ENTITLEMENT, DEPT FROM Timetables ORDER BY LN ASC;"

	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
	%>
	<div style="padding-left: 15px; font-size: 10pt;">
		Which Poor Soul Has Had To Cover Lots Of Classes? Find Out Here!<br>
		By Looking At This Report, You Will Be Able To Find Out Who Has Been Overloaded With PleaseTakes.
		You'll Then Be Able To Make Choices As To Who Will Cover Classes For The Future.<br>
		<b>This Report DOES NOT Include Outside Cover Staff.</b><br>
		Staff Who Have Had Their Entitlement Used Up Are Listed In <b>Bold</b>.
		<hr size="1">
		<div style="text-align: center; font-weight: bold;">Listed Staff Members : <i><%=RSUSER.RECORDCOUNT%></i></div>
		<br>
		<table class="m_l_sel_t" style="color: #000;" cellpadding="0" cellspacing="0">
			<tr>
				<td style="width: 50%; border-bottom: 1px solid #000; padding-left: 5px; font-size: 16pt;"><b>Staff Member</b></td>
				<td style="width: 25%; border-bottom: 1px solid #000; text-align: center; font-size: 16pt;"><b>Total Ent.</b></td>
				<td style="width: 25%; border-bottom: 1px solid #000; text-align: center; font-size: 16pt;"><b>Free Ent.</b></td>
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
			<tr>
				<td style="width: 50%; border-bottom: 1px solid #000; padding-left: 5px;<%if (RSUSER("ENTITLEMENT") - RSCOVER.RECORDCOUNT = 0) and (RSUSER("ENTITLEMENT") <> 0) then%> font-weight: bold;<%else%><%end if%>"><%=RSUSER("TITLE")%>. <%=left(RSUSER("FN"),1)%>. <%=RSUSER("LN")%>, <span style="font-size: 10pt;"><%=RSDEPT("FULL")%></span></td>
				<td style="width: 25%; border-bottom: 1px solid #000; text-align: center;<%if RSUSER("ENTITLEMENT") - RSCOVER.RECORDCOUNT = 0 and (RSUSER("ENTITLEMENT") <> 0) then%> font-weight: bold;<%else%><%end if%>"><%=RSUSER("ENTITLEMENT")%></td>
				<td style="width: 25%; border-bottom: 1px solid #000; text-align: center;<%if RSUSER("ENTITLEMENT") - RSCOVER.RECORDCOUNT = 0 and (RSUSER("ENTITLEMENT") <> 0) then%> font-weight: bold;<%else%><%end if%>"><%=RSUSER("ENTITLEMENT") - RSCOVER.RECORDCOUNT%></td>
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
		</table>
		<%
		RSUSER.close
		set RSUSER = nothing
		%>
	</div>
</div>
<div style="page-break-after: always;">
	<div style="font-size: 22pt; font-weight: bold; letter-spacing: -2px;">Number Crunching</div>
	<hr size="1">
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
	<div style="padding-left: 15px;">
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
	for d = 1 to 7
	if (d = 1) or (d = 7) then
		 if var_est_enabled_weekends = "0" then
		 else
%>
<div style="page-break-after: always;">
	<div style="font-size: 22pt; font-weight: bold; letter-spacing: -2px;">Detailed Cover Summary For <%=getdow((DateAdd("d",d - 1,getweekstart(date()))))%>, <%=DateAdd("d",d - 1,getweekstart(date()))%></div>
	<hr size="1">
	<div style="padding-left: 15px;">
		Below Is A Table Showing Everyone Who Was Listed As Absent On <%=getdow((DateAdd("d",d - 1,getweekstart(date()))))%>, And Who Covered Them In The Relevant Boxes. 
		<hr size="1">
		<%
		tempdaydow = daydow
		tempdaydate = daydate
		
		daydow = d
		daydate = DateAdd("d",d - 1,getweekstart(date()))
		%>
		<!--#include virtual="/pt/modules/ss/timetables/admin_rep_sum_p.inc"-->
		<%
		daydow = tempdaydow
		daydate = tempdaydate
		
		tempdaydow = null
		tempdaydate = null
		%>
	</div>
</div>
<%
	 end if
else	
%>
<div style="page-break-after: always;">
	<div style="font-size: 22pt; font-weight: bold; letter-spacing: -2px;">Detailed Cover Summary For <%=getdow((DateAdd("d",d - 1,getweekstart(date()))))%>, <%=DateAdd("d",d - 1,getweekstart(date()))%></div>
	<hr size="1">
	<div style="padding-left: 15px;">
		Below Is A Table Showing Everyone Who Was Listed As Absent On <%=getdow((DateAdd("d",d - 1,getweekstart(date()))))%>, And Who Covered Them In The Relevant Boxes. 
		<hr size="1">
		<%
		tempdaydow = daydow
		tempdaydate = daydate
		
		daydow = d
		daydate = DateAdd("d",d - 1,getweekstart(date()))
		%>
		<!--#include virtual="/pt/modules/ss/timetables/admin_rep_sum_p.inc"-->
		<%
		daydow = tempdaydow
		daydate = tempdaydate
		
		tempdaydow = null
		tempdaydate = null
		%>
	</div>
</div>
<%
		end if
	next


elseif pagetype = "5" then
	backupyear = left(request("backup"),4)
	backupweek = mid(request("backup"),6)

	RSBACKEDUPSQL = "SELECT * FROM Inventory WHERE weekno = " & backupweek & " AND backupyear = " & backupyear

	Set RSBACKEDUP = Server.CreateObject("Adodb.RecordSet")
	RSBACKEDUP.Open RSBACKEDUPSQL, backupconn, adopenkeyset, adlockoptimistic
%>
<table style="height: 1028px; width: 100%;" cellpadding="0" cellspacing="0">
	<tr>
		<td style="vertical-align: middle; text-align: right;">
			<img src="/pt/media/login/<%=var_est_logimg%>" border="0" alt="<%=var_est_full%>"><br>
			<img src="/pt/media/admin/p_wr.gif" border="0" alt="The Weekly Report"><br>
			<span style="color: #000; font-size: 18pt; font-weight: bold; letter-spacing: -2px;">
				For <%=var_est_full%><br>
				Week <%=RSBACKEDUP("StartDate")%> - <%=RSBACKEDUP("EndDate")%><br>
				Using PleaseTakes <%=var_ver%><br><br>
				BACKUP COPY
			</span>
		</td>
	</tr>
</table>
<%
	RSBACKEDUP.close
	set RSBACKEDUP = nothing

elseif pagetype = "6" then

	RSUSERSQL = "SELECT ID, LN, FN, TITLE, DEPT, CATEGORY, ENTITLEMENT, DEFROOM FROM Timetables ORDER BY ENTITLEMENT DESC, LN"

	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<div style="font-size: 22pt; font-weight: bold; letter-spacing: -2px;">PleaseTakes <%=var_ver%></div>
	<hr size="1">
	<div style="font-size: 12pt; padding-left: 10pt; font-weight: bold;">
		Staff Entitlement List<br>
		<span style="font-size: 10pt;">Staff Are Ordered In Order Of Number Of PleaseTakes/Week, Alphabetically. Printing <%=RSUSER.RECORDCOUNT%> Members</span>
	</div>
	<hr size="1">
	<div style="font-size: 12pt; padding-left: 10pt; font-weight: bold;">
	<%
		do until RSUSER.EOF
			RSDEPTSQL = "SELECT * FROM DEPARTMENTS WHERE DEPTID = " & RSUSER("DEPT")

			Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
			RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic
	%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_t" style="width: 80px; color: #000;"><b><%=RSUSER("ENTITLEMENT")%></b> <%if RSUSER("ENTITLEMENT") = "1" then%>Period<%else%>Periods<%end if%></td>
					<td class="m_l_list_t" style="color: #000;"><b><%=RSUSER("LN")%>, <%=Left(RSUSER("FN"),1)%>.</b> :: 
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
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
					RSDEPT.close
					set RSDEPT = nothing
				RSUSER.MOVENEXT
				loop
				%>
			</table>
		</div>
	</div>
	<%
				RSUSER.close
				set RSUSER = nothing
	%>
	</div>
<%
elseif pagetype = "7" then
	if (request("uid") = "") then

	RSCHECKSQL = "SELECT ID, LN, FN, TITLE FROM Timetables ORDER BY LN ASC"

	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<div style="font-size: 18pt; font-weight: bold; letter-spacing: -2px;">PleaseTakes <%=var_ver%></div>
	<hr size="1">
	<div style="font-size: 10pt; padding-left: 10pt; font-weight: bold;">
		Timetables For All Staff Members <i>(Total: <%=RSCHECK.RECORDCOUNT%>)</i>
	</div>
	<hr size="1">
<%
	if RSCHECK.RECORDCOUNT = 0 then
%>
	No Staff Were Found!
<%
	else
		nopages = round((RSCHECK.RECORDCOUNT / 2))
		leftover = RSCHECK.RECORDCOUNT mod 1
			
		if leftover <> 0 then
			nopages = nopages + 1
		else
			nopages = nopages
		end if
		
		for p = 1 to nopages
		
		if p = nopages then
		%>
		<div>
		<%
		else
		%>
		<div style="page-break-after: always;">
		<%
		end if
			for tt = 1 to 2
				if RSCHECK.EOF then
				else
		%>
					<%
					if (p = 1) and (tt = 1) then
					else
					%>
					<hr size="1">
					<%
					end if
					%>
					<div style="width: 100%; padding-bottom: 5px; font-size: 14pt; font-weight: bold; letter-spacing: -1px;"><%=RSCHECK("LN")%>, <%=left(RSCHECK("FN"),1)%>.</div>
					<%
					if var_est_enabled_weekends = "1" then
					%>
					<!--#include virtual="/pt/modules/ss/timetables/admin_rep_staff_7day_p.inc"-->
					<%
					else
					%>
					<!--#include virtual="/pt/modules/ss/timetables/admin_rep_staff_5day_p.inc"-->
					<%
					end if

					RSCHECK.MOVENEXT
				end if
			next
		%>
		</div>
		<%
		next
					

	RSCHECK.close
	set RSCHECK = nothing

	end if

	else
	
	if (isNumeric(request("uid")) = False) or (request("uid") = "") then
%>
	<div style="font-size: 18pt; font-weight: bold; letter-spacing: -2px;">PleaseTakes <%=var_ver%></div>
	<hr size="1">
	<div style="font-size: 10pt; padding-left: 10pt; font-weight: bold;">
		Individual Timetable Report
	</div>
	<hr size="1">
	An Error Has Occured, Please Try Again!
<%

	else
	
	RSCHECKSQL = "SELECT ID, LN, FN, TITLE FROM Timetables WHERE ID = " & request("uid")

	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<div style="font-size: 18pt; font-weight: bold; letter-spacing: -2px;">PleaseTakes <%=var_ver%></div>
	<hr size="1">
	<div style="font-size: 10pt; padding-left: 10pt; font-weight: bold;">
		Individual Timetable Report
	</div>
	<hr size="1">
	<div style="width: 100%; padding-bottom: 5px; font-size: 14pt; font-weight: bold; letter-spacing: -1px;"><%=RSCHECK("LN")%>, <%=left(RSCHECK("FN"),1)%>.</div>
	<%
	if var_est_enabled_weekends = "1" then
	%>
	<!--#include virtual="/pt/modules/ss/timetables/admin_rep_staff_7day_p.inc"-->
	<%
	else
	%>
	<!--#include virtual="/pt/modules/ss/timetables/admin_rep_staff_5day_p.inc"-->
	<%
	end if

	RSCHECK.close
	set RSCHECK = nothing
	
	end if

	end if

elseif pagetype = "8" then
	daydow = request("daydow")
	daydate = request("daydate")

			RSSLIPSSQL = "SELECT * FROM Cover WHERE ID = " & request("cover")
		
			Set RSSLIPS = Server.CreateObject("Adodb.RecordSet")
			RSSLIPS.Open RSSLIPSSQL, dataconn, adopenkeyset, adlockoptimistic
			
			if RSSLIPS("OCOVER") = "1" then
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
		
				RSCOVERINGSQL = "SELECT * FROM OCover WHERE ID = " & RSSLIPS("COVERING")
				Set RSCOVERING = Server.CreateObject("Adodb.RecordSet")
				RSCOVERING.Open RSCOVERINGSQL, dataconn, adopenkeyset, adlockoptimistic
			else
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
			end if
%>
	<div style="font-size: 18pt; font-weight: bold; letter-spacing: -2px;">PleaseTakes <%=var_ver%></div>
	<hr size="1">
	<div style="font-size: 10pt; padding-left: 10pt; font-weight: bold;">
		Individual PleaseTake Slip For <%=RSFOR("FN")%>&nbsp;<%=RSFOR("LN")%>, Printed At <%=date%>, <%=time%>
	</div>
	<hr size="1">
		<div>
			<div class="rep_slip_p">
				<div class="rep_slip_c_p">
					<div style="height: 22px; line-height: 15px; border-bottom: 1px solid #000;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr>
								<td style="width: 50%; color: #000; font-size: 10pt; font-family: Tahoma,Sans-Serif;"><%=getdow(daydow)%>, <%=day(daydate)%>&nbsp;<%=cal_getmonth(month(daydate))%>&nbsp;<%=year(daydate)%></td>
								<td style="width: 50%; color: #000; font-size: 10pt; text-align: right; font-family: Tahoma, Sans-Serif;"><b><%=var_pname%></b> <%=var_ver%></td>
							</tr>
						</table>
					</div>
					<div class="rep_slip_b_p">
					<table class="m_l_sel_t" style="height: 100%;" cellpadding="0" cellspacing="0">
						<tr>
							<td style="width: 60%;">
							<table class="m_l_sel_t" style="height: 100%;" cellpadding="0" cellspacing="0">
								<tr>
									<td colspan="2" class="rep_slip_t_p" style="border-bottom: 1px solid #000;"><b>Please Arrange To Take This Class.</b></td>
								</tr>
								<tr>
									<td class="rep_slip_t_p" style="height: 30px; width: 25%; border-bottom: 1px solid #000;">&nbsp;<b>Covering:&nbsp;&nbsp;</b></td>
									<td class="rep_slip_t_p" style="border-bottom: 1px solid #000;"><b><%=left(RSCOVERING("FN"),1)%>.&nbsp;<%=RSCOVERING("LN")%></b> <%if RSSLIPS("OCover") = "1" then%>(Outside Cover)<%else%>(<%=RSCOVERINGDEPT("SHORT")%><%if RSCOVERING("DEFROOM") <> "" then%>, Room <b><%=RSCOVERING("DEFROOM")%></b><%else%><%end if%>)<%end if%></td>
								</tr>
								<tr>
									<td class="rep_slip_t_p" style="height: 30px; width: 25%; border-bottom: 1px solid #000;">&nbsp;<b>For:</b></td>
									<td class="rep_slip_t_p" style="border-bottom: 1px solid #000;"><%=left(RSFOR("FN"),1)%>.&nbsp;<%=RSFOR("LN")%> (<%=RSFORDEPT("SHORT")%>)</td>
								</tr>
								<tr>
									<td class="rep_slip_t_p" style="height: 30px; width: 25%; border-bottom: 1px solid #000;">&nbsp;<b>Class:</b></td>
									<td class="rep_slip_t_p" style="border-bottom: 1px solid #000;"><%=RSCLASS(RSSLIPS("PERIOD") & "_" & daydow)%></td>
								</tr>
								<tr>
									<td colspan="2" class="rep_slip_t_p" style="text-align: right;"><b>Thank You Very Much!</b>&nbsp;</td>
								</tr>
							</table>
							</td>
							<td style="width: 40%; text-align: right;">
							<table class="m_l_sel_t" style="height: 100%;" cellpadding="0" cellspacing="0">
								<tr>
									<td class="rep_slip_t_p" style="width: 50%; height: 27px; border-bottom: 1px solid #000; border-left: 1px solid #000; border-right: 1px solid #000; text-align: center;"><b>Period</b></td>
									<td class="rep_slip_t_p" style="width: 50%; height: 27px; border-bottom: 1px solid #000; text-align: center;"><b>Room</b></td>
								</tr>
								<tr>
									<td class="rep_slip_t_p" style="border-left: 1px solid #000; border-right: 1px solid #000; text-align: center; font-size: 70pt;"><%=RSSLIPS("PERIOD")%></td>
									<td class="rep_slip_t_p" style="text-align: center; font-size: 36pt; letter-spacing: -6px;"><%=RSCLASS("R" & RSSLIPS("PERIOD") & "_" & daydow)%></td>
								</tr>
							</table>
							</td>
						</tr>
					</table>
					</div>
				</div>
			</div>
			<hr size="1">
		</div>
<%
	if RSSLIPS("OCover") = "1" then
	else
	RSCOVERINGDEPT.close
	end if
	RSCOVERING.close
	RSFORDEPT.close
	RSFOR.close
	RSCLASS.close
	RSCOVER.close
	if RSSLIPS("OCover") = "1" then
	else
	set RSCOVERINGDEPT = nothing
	end if
	set RSCOVERING = nothing
	set RSFORDEPT = nothing
	set RSFOR = nothing
	set RSCLASS = nothing
	set RSCOVER = nothing
	RSSLIPS.close
	set RSSLIPS = nothing
end if
%>
</div>

</body>

</html>
<%
end if
%><!--#include virtual="/pt/modules/ss/p_e.inc"-->