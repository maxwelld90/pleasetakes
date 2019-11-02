<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" >
<!--#include virtual="/pt/modules/ss/usersys/logincheck_std.inc"-->

<!--#include virtual="/pt/modules/ss/p_s.inc"-->
<%
pagetype = request("id")
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="../modules/css/std.css">
<script language="javascript" type="text/javascript" src="/pt/modules/js/std.js"></script>
<title><%=var_ptitle%></title>
</head>

<body>

<div class="smlb_b"></div>
<div class="topb_b"></div>

<div class="main">
	<!--#include virtual="/pt/modules/ss/topbar/std.inc"-->
<%
if pagetype = "2" then

	count = 0

	RSCHECKSQL = "SELECT * FROM Cover WHERE DAY = " & DOW & " AND COVERING = " & session("sess_ttid")
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<div class="m_l">
		<div class="m_l_title">View My PleaseTakes</div>
		<div class="m_l_subtitle">View PleaseTakes For Today, And Yesterday...</div>
		<div class="m_l_ins">...And The Day Before Yesterday...</div>
		<div class="m_l_sel">
			<%
			if RSCHECK.RECORDCOUNT = 0 then
			%>
			<div class="m_l_middletitle">No Cover Required From You!</div>
			<div style="padding: 4px 0px 8px 4px;">
			You Do Not Need To Cover Any Classes Today! Yipee!
			<%
			else
			%>
			<div class="m_l_middletitle">Your PleaseTakes For Today:</div>
			<div style="padding: 4px 0px 8px 4px;">
			<%
				do until RSCHECK.EOF
				
				RSCHECK2SQL = "SELECT * FROM Timetables WHERE ID = " & RSCHECK("For")
				Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
				RSCHECK2.Open RSCHECK2SQL, dataconn, adopenkeyset, adlockoptimistic
			%>
			<table class="m_l_sel_t" style="width: 546px;" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b">
					<td class="m_l_list_t rep_week_t" style="width: 10%; background-color: #A5C0E2;"><b>Period</b></td>
					<td class="m_l_list_t rep_week_t" style="width: 20%; background-color: #A5C0E2;"><b>Room</b></td>
					<td class="m_l_list_t rep_week_t" style="width: 70%; text-align: left; background-color: #A5C0E2;"><b>Covering For</b></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" style="height: 5px;" colspan="3"></td>
				</tr>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);">
					<td class="m_l_list_t rep_week_t" style="width: 10%;"><b><%=RSCHECK("Period")%></b></td>
					<td class="m_l_list_t rep_week_t" style="width: 20%;"><b><%=RSCHECK2("R" & RSCHECK("PERIOD") & "_" & DOW)%></b></td>
					<td class="m_l_list_t rep_week_t" style="width: 70%; text-align: left;"><%=RSCHECK2("FN")%>&nbsp;<%=RSCHECK2("LN")%></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" style="height: 5px;" colspan="3"></td>
				</tr>
			</table>
			<%
				count = count + 1
			
				RSCHECK2.close
				set RSCHECK2 = nothing
				RSCHECK.MOVENEXT
				loop			
			end if

			RSCHECK.close
			set RSCHECK = nothing
			%>
			</div>

			<div class="m_l_middletitle">Previous PleaseTakes</div>
			<div style="padding: 4px 0px 8px 0px;">
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('this');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_clock.gif" border="0" alt="PleaseTakes From This Current Week"></td>
					<td class="m_l_list_t"><b>From This Week</b></td>
				</tr>
				<%
				RSCHECKSQL = "SELECT * FROM Cover WHERE DAY < " & DOW & " AND COVERING = " & session("sess_ttid")
				Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
				RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
				%>
				<tr id="list_this" style="display: none;">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="line-height: 20px; padding-top: 7px;">
						<%
						if RSCHECK.RECORDCOUNT = 0 then
						%>
						No PleaseTakes Found!
						<%
						else
						
							do until RSCHECK.EOF

							RSCHECK2SQL = "SELECT * FROM Timetables WHERE ID = " & RSCHECK("For")
							Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
							RSCHECK2.Open RSCHECK2SQL, dataconn, adopenkeyset, adlockoptimistic
						%>
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr class="m_l_list_b">
								<td class="m_l_list_t rep_week_t" style="width: 10%; background-color: #A5C0E2;"><b>Day</b></td>
								<td class="m_l_list_t rep_week_t" style="width: 10%; background-color: #A5C0E2;"><b>Period</b></td>
								<td class="m_l_list_t rep_week_t" style="width: 20%; background-color: #A5C0E2;"><b>Room</b></td>
								<td class="m_l_list_t rep_week_t" style="width: 70%; text-align: left; background-color: #A5C0E2;"><b>Covering For</b></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="3"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);">
								<td class="m_l_list_t rep_week_t" style="width: 15%;"><%=getDOW(RSCHECK("DAY"))%></td>
								<td class="m_l_list_t rep_week_t" style="width: 10%;"><%=RSCHECK("PERIOD")%></td>
								<td class="m_l_list_t rep_week_t" style="width: 25%;"><%=RSCHECK2("R" & RSCHECK("PERIOD") & "_" & DOW)%></td>
								<td class="m_l_list_t rep_week_t" style="width: 50%; text-align: left;"><%=RSCHECK2("FN")%>&nbsp;<%=RSCHECK2("LN")%></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="3"></td>
							</tr>
						</table>
						<%
							count = count + 1

							RSCHECK.MOVENEXT
							loop
						end if
						%>
					</td>
				</tr>
				<%
				RSCHECK.close
				set RSCHECK = nothing
				%>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
				RSCHECKSQL = "SELECT * FROM Inventory Order By WeekNo AND BackupYear ASC"
				Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
				RSCHECK.Open RSCHECKSQL, backupconn, adopenkeyset, adlockoptimistic
				
				do until RSCHECK.EOF
				%>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('<%=RSCHECK("WEEKNO")%>_<%=RSCHECK("BACKUPYEAR")%>');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_clock.gif" border="0" alt="Week Beginning <%=RSCHECK("STARTDATE")%>"></td>
					<td class="m_l_list_t">Week <b><%=RSCHECK("STARTDATE")%> - <%=RSCHECK("ENDDATE")%></b></td>
				</tr>
				<tr id="list_<%=RSCHECK("WEEKNO")%>_<%=RSCHECK("BACKUPYEAR")%>" style="display: none;">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="line-height: 20px; padding-top: 7px;">
						<%
						RSCHECK2SQL = "SELECT * FROM [C_" & RSCHECK("STARTDATE") & "_" & RSCHECK("ENDDATE") & "_" & RSCHECK("WEEKNO") &"] WHERE COVERING = " & session("sess_ttid") & " ORDER BY DAY"
						Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
						RSCHECK2.Open RSCHECK2SQL, backupconn, adopenkeyset, adlockoptimistic
						
						if RSCHECK2.RECORDCOUNT = 0 then
						%>
						No PleaseTakes Found!
						<%
						else
						
							do until RSCHECK2.EOF

							RSCHECK3SQL = "SELECT * FROM Timetables WHERE ID = " & RSCHECK2("FOR")
							Set RSCHECK3 = Server.CreateObject("Adodb.RecordSet")
							RSCHECK3.Open RSCHECK3SQL, dataconn, adopenkeyset, adlockoptimistic
						%>
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr class="m_l_list_b">
								<td class="m_l_list_t rep_week_t" style="width: 10%; background-color: #A5C0E2;"><b>Day</b></td>
								<td class="m_l_list_t rep_week_t" style="width: 10%; background-color: #A5C0E2;"><b>Period</b></td>
								<td class="m_l_list_t rep_week_t" style="width: 20%; background-color: #A5C0E2;"><b>Room</b></td>
								<td class="m_l_list_t rep_week_t" style="width: 70%; text-align: left; background-color: #A5C0E2;"><b>Covering For</b></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="3"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);">
								<td class="m_l_list_t rep_week_t" style="width: 15%;"><%=getDOW(RSCHECK2("DAY"))%></td>
								<td class="m_l_list_t rep_week_t" style="width: 10%;"><%=RSCHECK2("PERIOD")%></td>
								<td class="m_l_list_t rep_week_t" style="width: 25%;"><%=RSCHECK3("R" & RSCHECK2("PERIOD") & "_" & DOW)%></td>
								<td class="m_l_list_t rep_week_t" style="width: 50%; text-align: left;"><%=RSCHECK3("FN")%>&nbsp;<%=RSCHECK3("LN")%></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="3"></td>
							</tr>
						</table>
						<%
							count = count + 1

							RSCHECK3.close
							set RSCHECK3 = nothing

							RSCHECK2.MOVENEXT
							loop
						end if
						
						RSCHECK2.close
						set RSCHECK2 = nothing
						%>
					</td>
				</tr>
				<%
				RSCHECK.MOVENEXT
				loop
				
				RSCHECK.close
				set RSCHECK = nothing
				%>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
			</table>
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 10pt;">Total Number Of PleaseTakes - <b><%=count%></b></div>
			</div>
		</div>
	</div>
<%
elseif pagetype = "3" then
%>
	<div class="m_l">
		<div class="m_l_title">View My Timetable</div>
		<div class="m_l_subtitle">See What Classes You Have, And When.</div>
		<div class="m_l_ins">Your Current Timetable Is Displayed Below.</div>
		<div class="m_l_sel">
		<%
		if var_est_enabled_weekends = "1" then
		%>
		<!--#include virtual="/pt/modules/ss/timetables/std_7day.inc"-->
		<%
		else
		%>
		<!--#include virtual="/pt/modules/ss/timetables/std_5day.inc"-->
		<%
		end if
		%>
		<br>
		This Timetable Only Displays Your Teaching Periods, Not PleaseTake Periods.<br>
		<b>If Anything Is Incorrect, Please Contact The System Administrator.</b>
		</div>
	</div>
<%
else
%>
	<div class="m_l">
		<div class="m_l_title">View My Information</div>
		<div class="m_l_subtitle">See Your PleaseTakes Or Timetable.</div>
		<div class="m_l_ins">Please Select A Task From Below.</div>
		<div class="m_l_sel">
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_t_l"><a href="view.asp?id=2"><img src="/pt/media/icons/48_view.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="view.asp?id=2">View My PleaseTakes</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="view.asp?id=3"><img src="/pt/media/icons/48_cal.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="view.asp?id=3">View My Timetable</a></td>
				</tr>
			</table>
		</div>
	</div>
<%
end if
%>
	<div class="m_r">
		<div class="m_r2">
		<!--#include virtual="/pt/modules/ss/rbar/std.inc"-->
		</div>
	</div>
</div>

</body>

</html>

<!--#include virtual="/pt/modules/ss/p_e.inc"-->