<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" >
<!--#include virtual="/pt/modules/ss/usersys/logincheck_admin.inc"-->

<!--#include virtual="/pt/modules/ss/p_s.inc"-->
<%
if session("sess_un") = settingsXML.documentElement.childNodes.item(0).childNodes.item(4).getAttribute("firstloginacc") then
	response.redirect "setup.asp?err=2"

else

pagetype = request("id")
errtype = request("err")

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
<script language="javascript" type="text/javascript" src="/pt/modules/js/admin.js"></script>
<title><%=var_ptitle%></title>
</head>

<body>

<div class="smlb_b"></div>
<div class="topb_b"></div>

<div class="main">
	<!--#include virtual="/pt/modules/ss/topbar/admin.inc"-->
<%
if pagetype = "6" then

	if (request("type")) = "2" then
	
		if session("sess_adminlevel") <> "1" then
			response.redirect "/pt/admin/cover.asp?id=6&type=1"
		else
		%>
	<div class="m_l">
		<div class="m_l_title">Step 1</div>
		<div class="m_l_subtitle">Which Day Do You Want To Select?</div>
		<div class="m_l_ins">Please Select The Date Which You Wish To Arrange Cover For.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<b>Please Select A Date From Below.</b>
			You Can Change The Months By Selecting The Arrows At The Top Left And Right Of The Calendar.
			<br>
			<!--#include virtual="/pt/modules/ss/timetables/admin_cover_cal_all.inc"-->
			<div style="padding-top: 10px;width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="cover.asp?id=1">&lt;Back</a></div>
		</div>
	</div>
	<div class="m_r">
		<div class="m_r2">
		<!--#include virtual="/pt/modules/ss/rbar/admin.inc"-->
		</div>
	</div>
		<%
		end if
	else
	%>
	Sorry, This Feature Has The Multi-Day Addition Still To Be Applied! 
	<%
	end if

elseif pagetype = "2" then

	coverday = cdate(request("coverday"))
	daydow = request("dow")

	if (request("type")) = "2" then
	
		if session("sess_adminlevel") <> "1" then
			response.redirect "/pt/admin/cover.asp?id=2&type=1"
		else
		RSDEPTSQL = "SELECT * FROM Departments ORDER BY Short"

		Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
		RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<form name="edit" action="/pt/modules/ss/db/add.asp?addtype=2&amp;type=2&amp;coverday=<%=coverday%>&amp;dow=<%=daydow%>" method="post">
	<div class="m_l">
		<div class="m_l_title">Step 2</div>
		<div class="m_l_subtitle">Choosing Who Will Be Absent</div>
		<div class="m_l_ins">Please Check The People Who Will Be Off On <%=GetDOW(coverday)%>, <%=coverday%>.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<b>Please Check Who Will Be Off At <%=var_est_short%> For On The Day Above.</b> Staff Are Arranged By Department.
			You Can Collapse And Expand Departments By Clicking On Their Names.<br>
			Once You Are Done, Please Click "Next" Below.
			<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="cover.asp?id=6&amp;type=2&amp;coverday=<%=coverday%>&amp;dow=<%=daydow%>">&lt;Back</a> :: <a href="#" onmouseup="document.edit.submit();">Next></a></div>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
			</table>
				<%
				do until RSDEPT.EOF
					RSUSERSQL = "SELECT * FROM Timetables WHERE DEPT = " & RSDEPT("DEPTID") & " ORDER BY LN"

					Set RSUSER = Server.CreateObject("Adodb.RecordSet")
					RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic

				%>
			<div class="m_l_middletitle" style="width: 546px; _width: 100%; margin-bottom: 10px; font-size: 12pt;" onmouseup="showdetail('<%=RSDEPT("DEPTID")%>');"><%=RSDEPT("FULL")%></div>
			<div id="list_<%=RSDEPT("DEPTID")%>" style="padding: 0px 0px 10px 0px; margin-top: -10px;">
				<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
					<%
					if RSUSER.RECORDCOUNT = 0 then
					%>
					<tr>
						<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
					</tr>
					<tr>
						<td class="m_l_list_t" style="height: 28px; width: 25px; padding-left: 5px;"></td>
						<td class="m_l_list_t">No Staff Were Found In This Department.</td>
					</tr>
					<%
					else
					do until RSUSER.EOF
	
						RSCHECKEDSQL = "SELECT ID, DAYDATE FROM Attendance WHERE USER = " & RSUSER("ID") & " AND DAY = " & daydow & " AND DAYDATE = #" & SQLDate(coverday) & "#"
	
						Set RSCHECKED = Server.CreateObject("Adodb.RecordSet")
						RSCHECKED.Open RSCHECKEDSQL, dataconn, adopenkeyset, adlockoptimistic
					%>
					<tr>
						<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
					</tr>
					<tr onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);">
						<td class="m_l_list_t" style="height: 28px; width: 25px; padding-left: 5px;">
						<input type="checkbox" name="u_<%=RSUSER("ID")%>"<%if RSCHECKED.RECORDCOUNT = 1 then%> checked<%else%><%end if%>>
						</td>
						<td class="m_l_list_t"><b><%=RSUSER("LN")%>, <%=Left(RSUSER("FN"),1)%>.</b></td>
					</tr>
					<%
					RSUSER.MOVENEXT
					total = total + 1
					loop
					end if
					%>
				</table>
			</div>
				<%
				RSDEPT.MOVENEXT
				loop
				%>
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 10pt;">Displaying A Total Of <b><%=total%></b> Staff Members</div>
			<div style="width: 100%; padding: 5px 0px 10px 0px; text-align: right; font-size: 18pt; font-weight: bold;"><a href="cover.asp?id=6&amp;type=2&amp;coverday=<%=coverday%>&amp;dow=<%=daydow%>">&lt;Back</a> :: <a href="#" onmouseup="document.edit.submit();">Next></a></div>
		</div>
	</div>
	<div class="m_r">
		<div class="m_r2">
		<!--#include virtual="/pt/modules/ss/rbar/admin.inc"-->
		</div>
	</div>
	</form>
<%
		end if
	else
		if session("sess_dept") <> "" then

	RSUSERSQL = "SELECT ID, FN, LN FROM Timetables WHERE DEPT = " & session("sess_dept") & " ORDER BY LN"

	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<form name="edit" action="/pt/modules/ss/db/add.asp?addtype=2&amp;type=1&amp;dept=<%=session("sess_dept")%>" method="post">
	<div class="m_l">
		<div class="m_l_title">Step 1</div>
		<div class="m_l_subtitle">Choosing Who Will Be Off Today</div>
		<div class="m_l_ins">Please Check The People Who Will Be Off In Your Department Today.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<b>Please Check Who Will Be Off In Your Department Today.</b><br>
			Once You Are Done, Please Click "Next" Below.
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
				do until RSUSER.EOF

					RSCHECKEDSQL = "SELECT ID FROM Attendance WHERE DEPT = " & session("sess_dept") & " AND USER = " & RSUSER("ID") & " AND DAY = " & DOW

					Set RSCHECKED = Server.CreateObject("Adodb.RecordSet")
					RSCHECKED.Open RSCHECKEDSQL, dataconn, adopenkeyset, adlockoptimistic
				%>
				<tr onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_t" style="height: 28px; width: 25px; padding-left: 5px;">
					<input type="checkbox" name="u_<%=RSUSER("ID")%>"<%if RSCHECKED.RECORDCOUNT = 1 then%> checked<%else%><%end if%>>
					</td>
					<td class="m_l_list_t"><b><%=RSUSER("LN")%>, <%=Left(RSUSER("FN"),1)%>.</b></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
					RSCHECKED.close
					set RSCHECKED = nothing
				RSUSER.MOVENEXT
				loop
				%>
			</table>
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 10pt;">Displaying A Total Of <b><%=RSUSER.RECORDCOUNT%></b> Staff Members</div>
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="cover.asp?id=1">&lt;Back</a> :: <a href="#" onmouseup="document.edit.submit();">Next></a></div>
		</div>
	</div>
	<div class="m_r">
		<div class="m_r2">
		<!--#include virtual="/pt/modules/ss/rbar/admin.inc"-->
		</div>
	</div>
	</form>
<%
	RSUSER.close
	set RSUSER = nothing
			else
				if session("sess_adminlevel") <> "1" then
					response.redirect "/pt/admin/cover.asp?err=2"
				else
					response.redirect "/pt/admin/cover.asp?id=2&type=2"
				end if
			end if
	end if

elseif pagetype = "3" then

	coverday = request("coverday")
	daydow = request("dow")

	if (request("type")) = "2" then

		if session("sess_adminlevel") <> "1" then
			response.redirect "/pt/admin/cover.asp?id=3&type=1"
		else
%>
	<div class="m_l">
		<div class="m_l_title">Step 3</div>
		<div class="m_l_subtitle">Selecting When Staff Will Be Off</div>
		<div class="m_l_ins">Once You Are Done, Please Click 'Next'.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Now You Have To Select When The Checked Staff Members Will Be Absent. They Are Listed Below In Timetable Form. <span style="font-weight: bold; color: #498F32;">Green</span> Periods Mean They Are Teaching, And Clicking The Box Will Turn Them To "Absent" (<span style="font-weight: bold; color: #F00;">Red</span>), Meaning They Will Require Cover.
			<span style="font-weight: bold; color: #752B73;">Purple</span> Cells Mean They Do Not Have A Class, And Are Free To Take A Class. If They Will Be Absent, Just Click The Relevant Period And The Letter Will Change To "A", Meaning Absent. If Someone Will Be Off For The Entire Day, Just Click
			Their Name And All Their Periods Will Turn To Absent, Regardless If They Have A Class Or Not.
			<br><br>
			<!--#include virtual="/pt/modules/ss/timetables/admin_allcover.inc"-->
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="cover.asp?id=2&amp;type=2&amp;coverday=<%=coverday%>&amp;dow=<%=daydow%>">&lt;Back</a> :: <a href="cover.asp?id=4&amp;type=2&amp;coverday=<%=coverday%>&amp;dow=<%=daydow%>">Next></a></div>
		</div>
	</div>
<%
		end if
	else
		if session("sess_dept") <> "" then
%>
	<div class="m_l">
		<div class="m_l_title">Step 2</div>
		<div class="m_l_subtitle">Selecting When The Staff Will Be Off</div>
		<div class="m_l_ins">Once You Are Done, Please Click 'Next'.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Now You Have To Select When The Checked Staff Members Will Be Absent. They Are Listed Below In Timetable Form. <span style="font-weight: bold; color: #498F32;">Green</span> Periods Mean They Are Teaching, And Clicking The Box Will Turn Them To "Absent" (<span style="font-weight: bold; color: #F00;">Red</span>), Meaning They Will Require Cover.
			<span style="font-weight: bold; color: #752B73;">Purple</span> Cells Mean They Do Not Have A Class, And Are Free To Take A Class. If They Will Be Absent, Just Click The Relevant Period And The Letter Will Change To "A", Meaning Absent. If Someone Will Be Off For The Entire Day, Just Click
			Their Name And All Their Periods Will Turn To Absent, Regardless If They Have A Class Or Not.
			<br><br>
			<!--#include virtual="/pt/modules/ss/timetables/admin_deptcover.inc"-->
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="cover.asp?id=2&type=1">&lt;Back</a> :: <a href="cover.asp?id=4&type=1">Next></a></div>
		</div>
	</div>
<%
		else
				if session("sess_adminlevel") <> "1" then
					response.redirect "/pt/admin/cover.asp?err=2"
				else
					response.redirect "/pt/admin/cover.asp?id=3&type=2"
				end if
		end if
	end if

elseif pagetype = "4" then

	if (request("type")) = "2" then

		if session("sess_adminlevel") <> "1" then
			response.redirect "/pt/admin/cover.asp?id=4&type=1"
		else

	coverday = request("coverday")
	daydow = request("dow")
	noperiods = 0

	RSDAYSQL = "SELECT * FROM Periods WHERE ID = " & daydow

	Set RSDAY = Server.CreateObject("Adodb.RecordSet")
	RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic

	RSEVERYONESQL = "SELECT * FROM Timetables"
					
	Set RSEVERYONE = Server.CreateObject("Adodb.RecordSet")
	RSEVERYONE.Open RSEVERYONESQL, dataconn, adopenkeyset, adlockoptimistic
	
	RSABSENTSQL = "SELECT * FROM Attendance WHERE DAY = " & daydow & " AND DAYDATE = #" & SQLDate(coverday) & "#"

	Set RSABSENT = Server.CreateObject("Adodb.RecordSet")
	RSABSENT.Open RSABSENTSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<form name="edit" action="/pt/modules/ss/db/add.asp?addtype=3&amp;type=2&amp;coverday=<%=coverday%>&amp;dow=<%=daydow%>" method="post">
	<div class="m_l">
		<div class="m_l_title">Step 4</div>
		<div class="m_l_subtitle">Choosing Who To Cover</div>
		<div class="m_l_ins">Please Read The Instructions Below.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<%
			if (request("err") = "1") then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t"><b>You Gave Someone Two Or More PleaseTakes For The Same Period!</b></td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err") = "2") then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t"><b>Someones Entitlement Has Been Reached.</b></td>
				</tr>
			</table>
			<hr size="1">
			<%
			end if
			%>
			For This Final Step, Please Select Which Staff You Want To Cover The Absent Staff. Once You Are Done, Click "Update" For The Arrangements To Be Made.
			<br><b>If You Do Not Click Update, Your Cover Arrangements Won't Be Stored!</b>
			<b>The Time Taken To Save Your Arrangements Will Vary Depending On How Many Staff Are Listed. Once You Are Done, Click "Finish".</b><br>
			<span style="font-size: 12pt;">This Is The Updated Page - Click "Update" To Store Cover!</span>
			<div style="padding-top: 10px;">


				<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<%
				do until RSABSENT.EOF
				
					RSUSERSQL = "SELECT * FROM Timetables WHERE ID = " & RSABSENT("USER")
					
					Set RSUSER = Server.CreateObject("Adodb.RecordSet")
					RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
				%>
					<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('<%=RSUSER("ID")%>');">
						<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="<%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%>"></td>
						<td class="m_l_list_t"><b><%=RSUSER("LN")%>, <%=left(RSUSER("FN"),1)%>.</b></td>
					</tr>
					<tr id="list_<%=RSUSER("ID")%>">
						<td class="m_l_list_p"></td>
						<td class="m_l_list_t" style="line-height: 20px; padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
						<%
						for i=1 to RSDAY("TOTALS")
						if RSABSENT(i & "_" & daydow) <> "" AND RSUSER(i & "_" & daydow) <> "" then
						%>
								<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
									<td class="m_l_list_t" style="width: 80px; height: 28px; padding: 3px 0px 3px 5px;">Period <b><%=i%></b><br><span style="font-size: 8pt;">Class <b><%=RSUSER(i & "_" & daydow)%></b></span></td>
									<td>
										<select class="b_std" size="1" name="<%=RSUSER("ID")%>_<%=i%>">
												<option value="">Please Select...<%for l=1 to 36%>&nbsp;<%next%></option>
										<optgroup label="From Department">
										<%
										dstaffcount = 0
										
										RSD_USERSQL = "SELECT * FROM Timetables WHERE (((Timetables.DEPT)=" & RSUSER("DEPT") & ") AND ((Timetables.[" & i & "_" & daydow & "]) Is Null)) AND ID <> " & RSUSER("ID")
					
										Set RSD_USER = Server.CreateObject("Adodb.RecordSet")
										RSD_USER.Open RSD_USERSQL, dataconn, adopenkeyset, adlockoptimistic
										
										do until RSD_USER.EOF
											RSD_ABSENTSQL = "SELECT * FROM Attendance WHERE (((Attendance.USER)=" & RSD_USER("ID") & ") AND ((Attendance.[" & i & "_" & daydow & "]) <> Null));"

											Set RSD_ABSENT = Server.CreateObject("Adodb.RecordSet")
											RSD_ABSENT.Open RSD_ABSENTSQL, dataconn, adopenkeyset, adlockoptimistic
											
											if RSD_ABSENT.RECORDCOUNT => 1 then
											else
													RSD_ALREADYSQL = "SELECT * FROM Cover WHERE COVERING = " & RSD_USER("ID") & " AND DAY = " & daydow & " AND DAYDATE = #" & SQLDate(coverday) & "# AND PERIOD = " & i

													Set RSD_ALREADY = Server.CreateObject("Adodb.RecordSet")
													RSD_ALREADY.Open RSD_ALREADYSQL, dataconn, adopenkeyset, adlockoptimistic
													
													if RSD_ALREADY.RECORDCOUNT => 1 then
													else

														RSD_ALREADYSAMESQL = "SELECT * FROM Cover WHERE FOR = " & RSUSER("ID") & " AND COVERING = " & RSD_USER("ID") & " AND DAY = " & daydow & " AND DAYDATE = #" & SQLDate(coverday) & "#"

														Set RSD_ALREADYSAME = Server.CreateObject("Adodb.RecordSet")
														RSD_ALREADYSAME.Open RSD_ALREADYSAMESQL, dataconn, adopenkeyset, adlockoptimistic
														
														if RSD_ALREADYSAME.RECORDCOUNT => 1 then
															RSD_COUNTSQL = "SELECT * FROM Cover WHERE COVERING = " & RSD_USER("ID")

															Set RSD_COUNT = Server.CreateObject("Adodb.RecordSet")
															RSD_COUNT.Open RSD_COUNTSQL, dataconn, adopenkeyset, adlockoptimistic
															
															periodsleft = RSD_USER("ENTITLEMENT") - RSD_COUNT.RECORDCOUNT
															
															if periodsleft =< 0 then
															else
																dstaffcount = dstaffcount + 1
										%>
												<option id="sd_<%=RSUSER("ID")%>_<%=i%>_<%=RSD_USER("ID")%>" value="<%=RSD_USER("ID")%>"><%=RSD_USER("LN")%>, <%=left(RSD_USER("FN"),1)%>. (<%=periodsleft%> Left) *</option>
										<%
															end if
															
															RSD_COUNT.close
															set RSD_COUNT = nothing
														else
														
															RSD_COUNTSQL = "SELECT * FROM Cover WHERE COVERING = " & RSD_USER("ID")

															Set RSD_COUNT = Server.CreateObject("Adodb.RecordSet")
															RSD_COUNT.Open RSD_COUNTSQL, dataconn, adopenkeyset, adlockoptimistic
															
															periodsleft = RSD_USER("ENTITLEMENT") - RSD_COUNT.RECORDCOUNT
															
															if periodsleft =< 0 then
															else
																dstaffcount = dstaffcount + 1
										%>
												<option id="sd_<%=RSUSER("ID")%>_<%=i%>_<%=RSD_USER("ID")%>" value="<%=RSD_USER("ID")%>"><%=RSD_USER("LN")%>, <%=left(RSD_USER("FN"),1)%>. (<%=periodsleft%> Left)</option>
										<%
															end if
														end if
												end if
											end if
										
											RSD_ABSENT.close
											set RDS_ABSENT = nothing										
										
										RSD_USER.MOVENEXT
										loop
										
										RSD_USER.close
										set RSD_USER = nothing
										
											if dstaffcount = 0 then
										%>
											<option value="">No Staff Available</option>
										<%
											else
											end if
											
											dstaffcount = null
										%>
										</optgroup>
										<optgroup label="Rest Of <%=var_est_short%>">
										<%
										astaffcount = 0
										
										RSA_USERSQL = "SELECT * FROM Timetables WHERE (((Timetables.DEPT)<>" & RSUSER("DEPT") & ") AND ((Timetables.[" & i & "_" & daydow & "]) Is Null)) AND ID <> " & RSUSER("ID")
					
										Set RSA_USER = Server.CreateObject("Adodb.RecordSet")
										RSA_USER.Open RSA_USERSQL, dataconn, adopenkeyset, adlockoptimistic
										
										do until RSA_USER.EOF
											RSA_ABSENTSQL = "SELECT * FROM Attendance WHERE (((Attendance.USER)=" & RSA_USER("ID") & ") AND ((Attendance.[" & i & "_" & daydow & "]) <> Null));"

											Set RSA_ABSENT = Server.CreateObject("Adodb.RecordSet")
											RSA_ABSENT.Open RSA_ABSENTSQL, dataconn, adopenkeyset, adlockoptimistic
											
											if RSA_ABSENT.RECORDCOUNT => 1 then
											else
													RSA_ALREADYSQL = "SELECT * FROM Cover WHERE COVERING = " & RSA_USER("ID") & " AND DAY = " & daydow & " AND DAYDATE = #" & SQLDate(coverday) & "# AND PERIOD = " & i

													Set RSA_ALREADY = Server.CreateObject("Adodb.RecordSet")
													RSA_ALREADY.Open RSA_ALREADYSQL, dataconn, adopenkeyset, adlockoptimistic
													
													if RSA_ALREADY.RECORDCOUNT => 1 then
													else

														RSA_ALREADYSAMESQL = "SELECT * FROM Cover WHERE FOR = " & RSUSER("ID") & " AND COVERING = " & RSA_USER("ID") & " AND DAY = " & daydow & " AND DAYDATE = #" & SQLDate(coverday) & "#"

														Set RSA_ALREADYSAME = Server.CreateObject("Adodb.RecordSet")
														RSA_ALREADYSAME.Open RSA_ALREADYSAMESQL, dataconn, adopenkeyset, adlockoptimistic
														
														if RSA_ALREADYSAME.RECORDCOUNT => 1 then
															RSA_COUNTSQL = "SELECT * FROM Cover WHERE COVERING = " & RSA_USER("ID")

															Set RSA_COUNT = Server.CreateObject("Adodb.RecordSet")
															RSA_COUNT.Open RSA_COUNTSQL, dataconn, adopenkeyset, adlockoptimistic
															
															periodsleft = RSA_USER("ENTITLEMENT") - RSA_COUNT.RECORDCOUNT
															
															if periodsleft =< 0 then
															else
																astaffcount = astaffcount + 1
										%>
												<option id="s_<%=RSUSER("ID")%>_<%=i%>_<%=RSA_USER("ID")%>" value="<%=RSA_USER("ID")%>"><%=RSA_USER("LN")%>, <%=left(RSA_USER("FN"),1)%>. (<%=periodsleft%> Left) *</option>
										<%
															end if
															
															RSA_COUNT.close
															set RSA_COUNT = nothing
														else
														
															RSA_COUNTSQL = "SELECT * FROM Cover WHERE COVERING = " & RSA_USER("ID")

															Set RSA_COUNT = Server.CreateObject("Adodb.RecordSet")
															RSA_COUNT.Open RSA_COUNTSQL, dataconn, adopenkeyset, adlockoptimistic
															
															periodsleft = RSA_USER("ENTITLEMENT") - RSA_COUNT.RECORDCOUNT
															
															if periodsleft =< 0 then
															else
																astaffcount = astaffcount + 1
										%>
												<option id="s_<%=RSUSER("ID")%>_<%=i%>_<%=RSA_USER("ID")%>" value="<%=RSA_USER("ID")%>"><%=RSA_USER("LN")%>, <%=left(RSA_USER("FN"),1)%>. (<%=periodsleft%> Left)</option>
										<%
															end if
														end if
												end if
											end if
										
											RSA_ABSENT.close
											set RDS_ABSENT = nothing										
										
										RSA_USER.MOVENEXT
										loop
										
										RSA_USER.close
										set RSA_USER = nothing
											if astaffcount = 0 then
										%>
											<option value="">No Staff Available</option>
										<%
											else
											end if
											
											astaffcount = null
										%>
										</optgroup>
										<optgroup label="Outside Cover">
										<%
										ostaffcount = 0
										dayfield = "D_" & daydow
										
										RSOC_USERSQL = "SELECT * FROM OCover WHERE [" & dayfield & "] = 1"
					
										Set RSOC_USER = Server.CreateObject("Adodb.RecordSet")
										RSOC_USER.Open RSOC_USERSQL, dataconn, adopenkeyset, adlockoptimistic
										
										do until RSOC_USER.EOF
											RSALREADY_USERSQL = "SELECT * FROM Cover WHERE [OCover] = 1 AND [COVERING] = " & RSOC_USER("ID") & " AND [PERIOD] = " & i & " AND [DAY] = " & daydow & " AND [DAYDATE] = #" & SQLDate(coverday) & "#"

											Set RSALREADY_USER = Server.CreateObject("Adodb.RecordSet")
											RSALREADY_USER.Open RSALREADY_USERSQL, dataconn, adopenkeyset, adlockoptimistic
											
											if RSALREADY_USER.RECORDCOUNT => 1 then
											else
															RSO_COUNTSQL = "SELECT * FROM Cover WHERE COVERING = " & RSOC_USER("ID")

															Set RSO_COUNT = Server.CreateObject("Adodb.RecordSet")
															RSO_COUNT.Open RSO_COUNTSQL, dataconn, adopenkeyset, adlockoptimistic
															
															periodsleft = RSOC_USER("ENTITLEMENT") - RSO_COUNT.RECORDCOUNT
															
															if periodsleft =< 0 then
															else
																ostaffcount = ostaffcount + 1
										%>
												<option id="o_<%=RSUSER("ID")%>_<%=i%>_<%=RSOC_USER("ID")%>" value="o_<%=RSOC_USER("ID")%>"><%=RSOC_USER("LN")%>,&nbsp;<%=left(RSOC_USER("FN"),1)%>. (<%=periodsleft%> Left)</option>
										<%
															end if
											end if
										RSOC_USER.MOVENEXT
										loop
											if ostaffcount = 0 then
										%>
											<option value="">No Staff Available</option>
										<%
											else
											end if
											
											ostaffcount = null
										%>
										</optgroup>
										</select>
									</td>
									<%
									RSWHOSQL = "SELECT * FROM Cover WHERE DAY = " & daydow & " AND DAYDATE = #" & SQLDate(coverday) & "# AND PERIOD = " & i & " AND FOR = " & RSUSER("ID")

									Set RSWHO = Server.CreateObject("Adodb.RecordSet")
									RSWHO.Open RSWHOSQL, dataconn, adopenkeyset, adlockoptimistic

									if RSWHO.RECORDCOUNT => 1 then
									%>
									<%
										RSWHOCOVERSQL = "SELECT FN, LN FROM Timetables WHERE ID = " & RSWHO("COVERING")

										Set RSWHOCOVER = Server.CreateObject("Adodb.RecordSet")
										RSWHOCOVER.Open RSWHOCOVERSQL, dataconn, adopenkeyset, adlockoptimistic
										
										if (rswho("ocover") = "1") then
											RSWHOCOVER2SQL = "SELECT FN, LN FROM OCover WHERE ID = " & RSWHO("COVERING")
							
											Set RSWHOCOVER2 = Server.CreateObject("Adodb.RecordSet")
											RSWHOCOVER2.Open RSWHOCOVER2SQL, dataconn, adopenkeyset, adlockoptimistic
									%>
									<td style="padding-right: 5px; text-align: right; font-size: 8pt; color: #456AAF;"><b><%=RSWHOCOVER2("LN")%>,&nbsp;<%=left(RSWHOCOVER2("FN"),1)%>.</b> (<a href="/pt/modules/ss/db/delete.asp?deltype=2&amp;type=2&amp;id=<%=RSWHO("ID")%>&amp;coverday=<%=coverday%>&amp;dow=<%=daydow%>">Deselect</a>)</td>
									<%
										else
									%>
									<td style="padding-right: 5px; text-align: right; font-size: 8pt;"><b><%=RSWHOCOVER("LN")%>,&nbsp;<%=left(RSWHOCOVER("FN"),1)%>.</b> (<a href="/pt/modules/ss/db/delete.asp?deltype=2&amp;type=2&amp;id=<%=RSWHO("ID")%>&amp;coverday=<%=coverday%>&amp;dow=<%=daydow%>">Deselect</a>)</td>
									<%
										end if

									else
									%>
									<td style="padding-right: 5px; text-align: right; font-size: 10pt; color: #F00;"><b>No Cover Chosen!</b></td>
									<%
									end if
									%>
								</tr>
								<tr>
									<td class="m_l_sel_sep" colspan="2"></td>
								</tr>
						<%
						else
							noperiods = noperiods + 1
						end if

						next

						if noperiods = RSDAY("TOTALS") then
						%>
						<tr>
							<td class="m_l_list_t"><%=RSUSER("FN")%> Is Listed As Being Absent, But Does Not Require Any Cover.</td>
						</tr>
						<%
						else
						end if
						
						noperiods = 0
						%>
						</table>
						</td>
					</tr>
					<tr>
						<td class="m_l_sel_sep" colspan="2"></td>
					</tr>
				<%
					RSUSER.close
					set RSUSER = nothing

				RSABSENT.MOVENEXT
				loop
				%>
				</table>

























			</div>
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="cover.asp?id=3&amp;type=2&amp;coverday=<%=coverday%>&amp;dow=<%=daydow%>">&lt;Back</a> :: <a href="#" onmouseup="location.href='/pt/modules/ss/db/delete.asp?deltype=7&amp;coverday=<%=request("coverday")%>&amp;dow=<%=request("dow")%>&amp;type=2'">Clear Cover</a> :: <a href="#" onmouseup="document.edit.submit();">Update</a> :: <a href="cover.asp?id=5&type=2">Finish></a></div>
		</div>
	</div>
	</form>
<%
	RSDAY.close
	RSEVERYONE.close
	RSABSENT.close
	set RSDAY = nothing
	set RSEVERYONE = nothing
	set RSABSENT = nothing
	
		end if
	else
	
		if session("sess_dept") <> "" then

	RSDAYSQL = "SELECT * FROM Periods WHERE ID = " & daydow

	Set RSDAY = Server.CreateObject("Adodb.RecordSet")
	RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic

	RSEVERYONESQL = "SELECT * FROM Timetables"
					
	Set RSEVERYONE = Server.CreateObject("Adodb.RecordSet")
	RSEVERYONE.Open RSEVERYONESQL, dataconn, adopenkeyset, adlockoptimistic
	
	RSABSENTSQL = "SELECT * FROM Attendance WHERE DAY = " & daydow & " AND DEPT = " & session("sess_dept")

	Set RSABSENT = Server.CreateObject("Adodb.RecordSet")
	RSABSENT.Open RSABSENTSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<form name="edit" action="/pt/modules/ss/db/add.asp?addtype=3&amp;type=1" method="post">
	<div class="m_l">
		<div class="m_l_title">Step 3</div>
		<div class="m_l_subtitle">Choosing Who To Cover</div>
		<div class="m_l_ins">Please Read The Instructions Below.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Staff And The Periods You Selected As Absent Are Listed Below. To Select Cover, Then Just Select Someone From The Relevant Drop-Down Box.
			Staff Are Listed Firstly From The Same Department, Then The Rest, And Finally, Outside Cover. Once The Request Has Been Saved, The Staff Member Will Appear Beside The Drop-Down Box. If The Text Is Brown,
			Another Member Of Teaching Staff Has Been Selected. If It's <span style="color: #498F32;"><b>Green</b></span>, Then An Outside Cover Staff Member Has Been Selected. Finally, If It's <span style="color: #F00;"><b>Red</b></span>, Then
			There Is A Problem.
			<b>The Time Taken To Process Your Request Will Vary Depending On How Many Staff Are Listed. Once You Are Done, Click "Finish".</b>
			<div style="padding-top: 10px;">
				<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<%
				do until RSABSENT.EOF
				
					RSUSERSQL = "SELECT * FROM Timetables WHERE ID = " & RSABSENT("USER")
					
					Set RSUSER = Server.CreateObject("Adodb.RecordSet")
					RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
				%>
					<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('<%=RSUSER("ID")%>');">
						<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="<%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%>"></td>
						<td class="m_l_list_t"><b><%=RSUSER("LN")%>, <%=left(RSUSER("FN"),1)%>.</b></td>
					</tr>
					<tr id="list_<%=RSUSER("ID")%>">
						<td class="m_l_list_p"></td>
						<td class="m_l_list_t" style="line-height: 20px; padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
						<%
						for i=1 to RSDAY("TOTALS")
						if RSABSENT(i & "_" & DOW) <> "" AND RSUSER(i & "_" & DOW) <> "" then
						%>
								<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
									<td class="m_l_list_t" style="width: 100px; height: 28px; padding-left: 5px;">Period <b><%=i%></b> Cover:</td>
									<td class="m_l_list_t" style="width: 190px;">
										<select class="b_std" size="1" name="<%=RSUSER("ID")%>_<%=i%>" onchange="document.edit.submit();">
												<option value="">Please Select...</option>
												<optgroup label="From Department">
										<%
										RSD_USERSQL = "SELECT * FROM Timetables WHERE (((Timetables.DEPT)=" & RSUSER("DEPT") & ") AND ((Timetables.[" & i & "_" & DOW & "]) Is Null)) AND ID <> " & RSUSER("ID")
					
										Set RSD_USER = Server.CreateObject("Adodb.RecordSet")
										RSD_USER.Open RSD_USERSQL, dataconn, adopenkeyset, adlockoptimistic
										
										do until RSD_USER.EOF
											RSD_ABSENTSQL = "SELECT * FROM Attendance WHERE (((Attendance.USER)=" & RSD_USER("ID") & ") AND ((Attendance.[" & i & "_" & DOW & "]) <> Null));"

											Set RSD_ABSENT = Server.CreateObject("Adodb.RecordSet")
											RSD_ABSENT.Open RSD_ABSENTSQL, dataconn, adopenkeyset, adlockoptimistic
											
											if RSD_ABSENT.RECORDCOUNT => 1 then
											else
													RSD_ALREADYSQL = "SELECT * FROM Cover WHERE COVERING = " & RSD_USER("ID") & " AND DAY = " & DOW & " AND PERIOD = " & i

													Set RSD_ALREADY = Server.CreateObject("Adodb.RecordSet")
													RSD_ALREADY.Open RSD_ALREADYSQL, dataconn, adopenkeyset, adlockoptimistic
													
													if RSD_ALREADY.RECORDCOUNT => 1 then
													else

														RSD_ALREADYSAMESQL = "SELECT * FROM Cover WHERE FOR = " & RSUSER("ID") & " AND COVERING = " & RSD_USER("ID") & " AND DAY = " & DOW

														Set RSD_ALREADYSAME = Server.CreateObject("Adodb.RecordSet")
														RSD_ALREADYSAME.Open RSD_ALREADYSAMESQL, dataconn, adopenkeyset, adlockoptimistic
														
														if RSD_ALREADYSAME.RECORDCOUNT => 1 then
														else
														
															RSD_COUNTSQL = "SELECT * FROM Cover WHERE COVERING = " & RSD_USER("ID")

															Set RSD_COUNT = Server.CreateObject("Adodb.RecordSet")
															RSD_COUNT.Open RSD_COUNTSQL, dataconn, adopenkeyset, adlockoptimistic
															
															periodsleft = RSD_USER("ENTITLEMENT") - RSD_COUNT.RECORDCOUNT
															
															if periodsleft =< 0 then
															else
										%>
												<option id="sd_<%=RSUSER("ID")%>_<%=i%>_<%=RSD_USER("ID")%>" value="<%=RSD_USER("ID")%>"><%=RSD_USER("LN")%>, <%=left(RSD_USER("FN"),1)%>. (<%=periodsleft%> Left)</option>
										<%
															end if
														end if
												end if
											end if
										
											RSD_ABSENT.close
											set RDS_ABSENT = nothing										
										
										RSD_USER.MOVENEXT
										loop
										
										RSD_USER.close
										set RSD_USER = nothing
										%>
												</optgroup>
												<optgroup label="Rest Of <%=var_est_short%>">
										<%
										RSA_USERSQL = "SELECT * FROM Timetables WHERE (((Timetables.DEPT)<>" & RSUSER("DEPT") & ") AND ((Timetables.[" & i & "_" & DOW & "]) Is Null)) AND ID <> " & RSUSER("ID")
					
										Set RSA_USER = Server.CreateObject("Adodb.RecordSet")
										RSA_USER.Open RSA_USERSQL, dataconn, adopenkeyset, adlockoptimistic
										
										do until RSA_USER.EOF
											RSA_ABSENTSQL = "SELECT * FROM Attendance WHERE (((Attendance.USER)=" & RSA_USER("ID") & ") AND ((Attendance.[" & i & "_" & DOW & "]) <> Null));"

											Set RSA_ABSENT = Server.CreateObject("Adodb.RecordSet")
											RSA_ABSENT.Open RSA_ABSENTSQL, dataconn, adopenkeyset, adlockoptimistic
											
											if RSA_ABSENT.RECORDCOUNT => 1 then
											else
													RSA_ALREADYSQL = "SELECT * FROM Cover WHERE COVERING = " & RSA_USER("ID") & " AND DAY = " & DOW & " AND PERIOD = " & i

													Set RSA_ALREADY = Server.CreateObject("Adodb.RecordSet")
													RSA_ALREADY.Open RSA_ALREADYSQL, dataconn, adopenkeyset, adlockoptimistic
													
													if RSA_ALREADY.RECORDCOUNT => 1 then
													else

														RSA_ALREADYSAMESQL = "SELECT * FROM Cover WHERE FOR = " & RSUSER("ID") & " AND COVERING = " & RSA_USER("ID") & " AND DAY = " & DOW

														Set RSA_ALREADYSAME = Server.CreateObject("Adodb.RecordSet")
														RSA_ALREADYSAME.Open RSA_ALREADYSAMESQL, dataconn, adopenkeyset, adlockoptimistic
														
														if RSA_ALREADYSAME.RECORDCOUNT => 1 then
														else
														
															RSA_COUNTSQL = "SELECT * FROM Cover WHERE COVERING = " & RSA_USER("ID")

															Set RSA_COUNT = Server.CreateObject("Adodb.RecordSet")
															RSA_COUNT.Open RSA_COUNTSQL, dataconn, adopenkeyset, adlockoptimistic
															
															periodsleft = RSA_USER("ENTITLEMENT") - RSA_COUNT.RECORDCOUNT
															
															if periodsleft =< 0 then
															else
										%>
												<option id="s_<%=RSUSER("ID")%>_<%=i%>_<%=RSA_USER("ID")%>" value="<%=RSA_USER("ID")%>"><%=RSA_USER("LN")%>, <%=left(RSA_USER("FN"),1)%>. (<%=periodsleft%> Left)</option>
										<%
															end if
														end if
												end if
											end if
										
											RSA_ABSENT.close
											set RDS_ABSENT = nothing										
										
										RSA_USER.MOVENEXT
										loop
										
										RSA_USER.close
										set RSA_USER = nothing
										%>
												</optgroup>
												<optgroup label="Outside Cover">
										<%
										dayfield = "D_" & DOW
										
										RSOC_USERSQL = "SELECT * FROM OCover WHERE [" & dayfield & "] = 1"
					
										Set RSOC_USER = Server.CreateObject("Adodb.RecordSet")
										RSOC_USER.Open RSOC_USERSQL, dataconn, adopenkeyset, adlockoptimistic
										
										do until RSOC_USER.EOF
											RSALREADY_USERSQL = "SELECT * FROM Cover WHERE [OCover] = 1 AND [COVERING] = " & RSOC_USER("ID") & " AND [PERIOD] = " & i & " AND [DAY] = " & DOW

											Set RSALREADY_USER = Server.CreateObject("Adodb.RecordSet")
											RSALREADY_USER.Open RSALREADY_USERSQL, dataconn, adopenkeyset, adlockoptimistic
											
											if RSALREADY_USER.RECORDCOUNT => 1 then
											else
															RSO_COUNTSQL = "SELECT * FROM Cover WHERE COVERING = " & RSOC_USER("ID")

															Set RSO_COUNT = Server.CreateObject("Adodb.RecordSet")
															RSO_COUNT.Open RSO_COUNTSQL, dataconn, adopenkeyset, adlockoptimistic
															
															periodsleft = RSOC_USER("ENTITLEMENT") - RSO_COUNT.RECORDCOUNT
															
															if periodsleft =< 0 then
															else
										%>
												<option id="o_<%=RSUSER("ID")%>_<%=i%>_<%=RSOC_USER("ID")%>" value="o_<%=RSOC_USER("ID")%>"><%=RSOC_USER("LN")%>,&nbsp;<%=left(RSOC_USER("FN"),1)%>. (<%=periodsleft%> Left)</option>
										<%
															end if
											end if
										RSOC_USER.MOVENEXT
										loop
										%>
												</optgroup>
										</select>
									</td>
									<%
									RSWHOSQL = "SELECT * FROM Cover WHERE DAY = " & DOW & " AND PERIOD = " & i & " AND FOR = " & RSUSER("ID")

									Set RSWHO = Server.CreateObject("Adodb.RecordSet")
									RSWHO.Open RSWHOSQL, dataconn, adopenkeyset, adlockoptimistic

									if RSWHO.RECORDCOUNT => 1 then
										RSWHOCOVERSQL = "SELECT FN, LN FROM Timetables WHERE ID = " & RSWHO("COVERING")

										Set RSWHOCOVER = Server.CreateObject("Adodb.RecordSet")
										RSWHOCOVER.Open RSWHOCOVERSQL, dataconn, adopenkeyset, adlockoptimistic
										
										if RSWHOCOVER.RECORDCOUNT =< 0 then
											RSWHOCOVER2SQL = "SELECT FN, LN FROM OCover WHERE ID = " & RSWHO("COVERING")

											Set RSWHOCOVER2 = Server.CreateObject("Adodb.RecordSet")
											RSWHOCOVER2.Open RSWHOCOVER2SQL, dataconn, adopenkeyset, adlockoptimistic
											
											if RSWHOCOVER2.RECORDCOUNT => 1 then
									%>
									<td style="padding-left: 10px; font-size: 10pt; color: #498F32;">Selected: <b><%=RSWHOCOVER2("LN")%>,&nbsp;<%=left(RSWHOCOVER2("FN"),1)%>.</b> (<a href="/pt/modules/ss/db/delete.asp?deltype=2&amp;type=1&amp;id=<%=RSWHO("ID")%>">Deselect</a>)</td>
									<%
											else
									%>
									<td style="padding-left: 10px; font-size: 10pt; color: #F00;"><b>Unable To Locate Cover!</b></td>
									<%
											end if
										else
									%>
									<td style="padding-left: 10px; font-size: 10pt;">Selected: <b><%=RSWHOCOVER("LN")%>,&nbsp;<%=left(RSWHOCOVER("FN"),1)%>.</b> (<a href="/pt/modules/ss/db/delete.asp?deltype=2&amp;type=1&amp;id=<%=RSWHO("ID")%>">Deselect</a>)</td>
									<%
										end if
									else
									%>
									<td style="padding-left: 10px; font-size: 10pt; color: #F00;">No Cover Chosen!</td>
									<%
									end if
									%>
								</tr>
								<tr>
									<td class="m_l_sel_sep" colspan="2"></td>
								</tr>
						<%
						else
							noperiods = noperiods + 1
						end if

						next

						if noperiods = RSDAY("TOTALS") then
						%>
						<tr>
							<td class="m_l_list_t"><%=RSUSER("FN")%> Is Listed As Being Absent, But Does Not Require Any Cover.</td>
						</tr>
						<%
						else
						end if
						
						noperiods = 0
						%>
						</table>
						</td>
					</tr>
					<tr>
						<td class="m_l_sel_sep" colspan="2"></td>
					</tr>
				<%
					RSUSER.close
					set RSUSER = nothing

				RSABSENT.MOVENEXT
				loop
				%>
				</table>
			</div>
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="cover.asp?id=3&type=1">&lt;Back</a> :: <a href="cover.asp?id=5&type=1">Finish></a></div>
		</div>
	</div>
	</form>
<%
	RSDAY.close
	RSEVERYONE.close
	RSABSENT.close
	set RSDAY = nothing
	set RSEVERYONE = nothing
	set RSABSENT = nothing
	
		else
				if session("sess_adminlevel") <> "1" then
					response.redirect "/pt/admin/cover.asp?err=2"
				else
					response.redirect "/pt/admin/cover.asp?id=4&type=2"
				end if
		end if

	end if

elseif pagetype = "5" then

	if (request("type")) = "2" then

		if session("sess_adminlevel") <> "1" then
			response.redirect "/pt/admin/cover.asp?id=3&type=1"
		else
%>
	<div class="m_l">
		<div class="m_l_title">Congratulations!</div>
		<div class="m_l_subtitle">You Have Completed The Staff Cover Wizard!</div>
		<div class="m_l_ins">Your Staff Have Been Successfullly Covered.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Congratulations, <%=session("sess_fn")%>! You Have Completed The Staff Cover Wizard!<br>
			Any Additions/Changes Have Been Successfully Saved. You Can Now Print Today's <a href="reports.asp?id=2&amp;print=1">PleaseTake Slips</a>, If You Require.
			Remember, If You Need To Make Changes Later On, Just Come Back And The System Will Be Updated.
			<div class="botopts">
				<ul>
					<li><a href="reports.asp?id=2&amp;print=1">Print Today's PleaseTake Slips</a></li>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
		end if
	else
		if session("sess_dept") <> "" then
%>
	<div class="m_l">
		<div class="m_l_title">Congratulations!</div>
		<div class="m_l_subtitle">You Have Completed The Staff Cover Wizard!</div>
		<div class="m_l_ins">Your Staff Have Been Successfullly Covered.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Congratulations, <%=session("sess_fn")%>! You Have Successfully Completed The Departmental Cover Wizard.
			Your Additions/Changed Have Been Successfully Saved And Slips Can Now Be Printed By The System Administrator For Distribution.<br>
			Remember, You Can Always Revisit This Wizard Later On To Make Changes!
			<div class="botopts">
				<ul>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
		else
				if session("sess_adminlevel") <> "1" then
					response.redirect "/pt/admin/cover.asp?err=2"
				else
					response.redirect "/pt/admin/cover.asp?id=5&type=2"
				end if
		end if
	end if

elseif errtype = "1" then
%>
	<div class="m_l">
		<div class="m_l_title">Sorry!</div>
		<div class="m_l_subtitle">You Cannot Continue!</div>
		<div class="m_l_ins">Please Read What Is Below.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Sorry, <%=session("sess_fn")%>, But The Wizard Is Unable To Continue As You Have Not Selected Any Staff That Will Be Off For Today.<br>
			Please Return To The Previous Page And Select Some Staff That Will Be Off.<br>
			<span style="font-size: 14pt; font-weight: bold;">If No Staff Will Be Off For The Selected Day, You Do Not Need To Use This Wizard.</span>
			<div class="botopts">
				<ul>
					<li><a href="#" onmouseup="history.back();">Return To The Previous Page</a></li>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
elseif errtype = "2" then
%>
	<div class="m_l">
		<div class="m_l_title">Sorry!</div>
		<div class="m_l_subtitle">The Staff Cover Wizard Can't Be Run!</div>
		<div class="m_l_ins">Your Account Has No Department Specified!</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Sorry, <%=session("sess_fn")%>, But The Staff Cover Wizard Cannot Be Run Under Your Account.<br>
			You Have No Department Linked To Your Account, And Do Not Have The Rights To Run The Full Staff Cover Wizard.
			<div class="botopts">
				<ul>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
else
%>
	<div class="m_l">
		<div class="m_l_title">Welcome!</div>
		<div class="m_l_subtitle">Welcome To The Staff Cover Wizard!</div>
		<div class="m_l_ins">Please Read What Is Below, And Let's Begin!</div>
		<div class="m_l_sel" style="width: 100%;">
			<%
			if session("sess_adminlevel") = "1" then
				if session("sess_dept") <> "" then
			%>
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			With This Wizard, You Can Arrange Staff Cover For A Chosen Day.
			Split Into Four Easy Steps, You'll Have The PleaseTake Slips Ready To Print In No Time!<br><br>
			To Begin, Just Click "Arrange Cover" Below. If You Want To First Select Some Outside Cover Staff To Be Available, Click "Arrange Outside Cover". 
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="cover.asp?id=6&type=1">Your Department Only></a><br><a href="cover.asp?id=6&type=2">Everyone At <%=var_est_short%>></a><br><a href="ocover.asp?id=1">Arrange Outside Cover First></a></div>
			<%
				else
			%>
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			With This Wizard, You Can Arrange Staff Cover For A Chosen Day.<br>
			Split Into Four Easy Steps, You'll Have The PleaseTake Slips Ready To Print In No Time!<br><br>
			To Begin, Just Click "Arrange Cover" Below.
			If You Want To First Select Some Outside Cover Staff To Be Available, Click "Arrange Outside Cover".			
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;">
			<a href="cover.asp?id=6&type=2">Arrange Cover></a><br>
			<a href="ocover.asp?id=1">Arrange Outside Cover></a>
			</div>
			<%
				end if
			else
				if session("sess_dept") <> "" then
			%>
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			This Wizard Has Been Designed To Be As Easy To Use As Possible, By Simply Using Only <b>Relevant</b> Information.<p>
			What You Will Do In This Wizard Is Choose Who In Your Department Is Off, And Select Another Member Of Staff To Cover For Them At Each Period.<p>
			To Begin, Click "Next" Below.
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="cover.asp?id=2&type=1">Next></a></div>
			<%
				else
					response.redirect "/pt/admin/cover.asp?err=2"
				end if
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