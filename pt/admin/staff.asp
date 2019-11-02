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
<script language="javascript" type="text/javascript" src="/pt/modules/js/admin.js"></script>
<title><%=var_ptitle%></title>
</head>

<body<%if (pagetype = "6") AND ((request("err")) = "") AND ((request("good")) = "") then%> onload="document.add.reset();"<%else%><%end if%>>

<div class="smlb_b"></div>
<div class="topb_b"></div>

<div class="main">
	<!--#include virtual="/pt/modules/ss/topbar/admin.inc"-->

<%
if pagetype = "2" then
%>
	<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->
<%

	if (request("order")) = "LN" then
	RSUSERSQL = "SELECT ID, LN, FN, TITLE, DEPT, CATEGORY, ENTITLEMENT, DEFROOM FROM Timetables ORDER BY LN"
	elseif (request("order")) = "FN" then
	RSUSERSQL = "SELECT ID, LN, FN, TITLE, DEPT, CATEGORY, ENTITLEMENT, DEFROOM FROM Timetables ORDER BY FN"
	elseif (request("order")) = "DEPT" then
	RSUSERSQL = "SELECT Departments.SHORT, Timetables.ID, Timetables.TITLE, Timetables.LN, Timetables.FN, Timetables.DEPT, Timetables.DEFROOM, Timetables.ENTITLEMENT, Timetables.CATEGORY FROM Departments INNER JOIN Timetables ON Departments.DEPTID = Timetables.DEPT ORDER BY Departments.SHORT, Timetables.LN"
	elseif (request("order")) = "ENT" then
	RSUSERSQL = "SELECT ID, LN, FN, TITLE, DEPT, CATEGORY, ENTITLEMENT, DEFROOM FROM Timetables ORDER BY ENTITLEMENT DESC"
	else
	RSUSERSQL = "SELECT ID, LN, FN, TITLE, DEPT, CATEGORY, ENTITLEMENT, DEFROOM FROM Timetables ORDER BY LN"
	end if
		
	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<div class="m_l">
		<div class="m_l_title">Staff Information List</div>
		<div class="m_l_subtitle">View Everyone's Details At One Glance.</div>
		<div class="m_l_ins">For More Information, Just Click A Staff Member.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<div style="width: 100%; text-align: center;">Order... <span class="m_l_order"><a href="staff.asp?id=2&amp;order=LN"><b>By Surname</b></a></span> :: <span class="m_l_order"><a href="staff.asp?id=2&amp;order=FN"><b>By First Name</b></a></span> :: <span class="m_l_order"><a href="staff.asp?id=2&amp;order=DEPT"><b>By Department</b></a></span> :: <span class="m_l_order"><a href="staff.asp?id=2&amp;order=ENT"><b>By Entitlement</b></a></span></div>
			<hr size="1">
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<%
				do until RSUSER.EOF
				
				RSDEPTSQL = "SELECT * FROM Departments WHERE DEPTID = " & RSUSER("DEPT")
					
				Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
				RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic
				%>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('<%=RSUSER("id")%>');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="<%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%>"></td>
					<td class="m_l_list_t"><b><%=RSUSER("LN")%>, <%=Left(RSUSER("FN"),1)%>.</b></td>
				</tr>
				<tr id="list_<%=RSUSER("id")%>" style="display: none;">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="line-height: 20px; padding-top: 7px;">
					<b>Full Name:</b> <%=RSUSER("Title")%>. <%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%> (System ID No. <%=RSUSER("ID")%>)<br>
					<b>Department:</b> <%=RSDEPT("FULL")%><br>
					<b>Position:</b>
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
					%><br>
					<b>Usual Room:</b> 
					<%
					if RSUSER("DEFROOM") <> "" then
					response.write RSUSER("DEFROOM")
					else
					response.write "N/A"
					end if
					%><br>
					<b>Period Entitlement:</b> <%=RSUSER("ENTITLEMENT")%> <%if RSUSER("ENTITLEMENT") = 1 then%> Period<%else%> Periods<%end if%>/Week<br>
					<%
					RSACCCHECKSQL = "SELECT * FROM Users WHERE TTID = " & RSUSER("ID")
					
					Set RSACCCHECK = Server.CreateObject("Adodb.RecordSet")
					RSACCCHECK.Open RSACCCHECKSQL, userconn, adopenkeyset, adlockoptimistic

					RSACCCHECK2SQL = "SELECT * FROM Admin WHERE TTID = " & RSUSER("ID")
					
					Set RSACCCHECK2 = Server.CreateObject("Adodb.RecordSet")
					RSACCCHECK2.Open RSACCCHECK2SQL, userconn, adopenkeyset, adlockoptimistic
					%>
					<b>PleaseTakes Account?</b> <%if RSACCCHECK.RECORDCOUNT => 1 then%>Yes, <b><%=RSACCCHECK("UN")%></b><%elseif RSACCCHECK2.RECORDCOUNT => 1 then%>Yes, <b><%=RSACCCHECK2("UN")%></b> (Level <%=RSACCCHECK2("ACCLEVEL")%> Admin Account)<%else%>No<%end if%>
					<hr size="1">
					<div style="width: 100%; text-align: center;"><b><%if (RSACCCHECK.RECORDCOUNT => 1) OR (RSACCCHECK2.RECORDCOUNT => 1) then%><a href="staff.asp?id=8&uid=<%=RSUSER("ID")%>&amp;manage=1">Manage Account</a><%else%><a href="staff.asp?id=8&amp;uid=<%=RSUSER("ID")%>">Create Account</a><%end if%> :: <a href="staff.asp?id=3&amp;user=<%=RSUSER("ID")%>">View/Edit Timetable</a></b> :: <b><a href="staff.asp?id=5&amp;user=<%=RSUSER("ID")%>">Edit Information</a></b> :: <b><a href="staff.asp?id=7&amp;confirm=1&amp;user=<%=RSUSER("ID")%>">Delete This Person</a></b></div>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
					RSACCCHECK.close
					set RSACCCHECK = nothing
					RSACCCHECK2.close
					set RSACCCHECK2 = nothing

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

elseif pagetype = "3" then

	if ((request("user")) <> "") then

	RSUSERSQL = "SELECT ID, FN, LN FROM Timetables WHERE id = " & request("user")
	
	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic

	if RSUSER.RECORDCOUNT = 0 then
%>
	<div class="m_l">
		<div class="m_l_title">Staff Timetable</div>
		<div class="m_l_subtitle">Cannot Find A User!</div>
		<div class="m_l_ins">The User ID Specified Is Invalid!</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Sorry, <%=session("sess_fn")%>, But We Cannot Find A User For ID <b>No.<%=request("user")%></b>, And Therefore We Do
			Not Have The Data Required To Draw A Timetable.<br>
			Please Choose An Option From Below To Continue.
			<div class="botopts">
				<ul>
					<li><a href="#" onmouseup="history.back();">Return</a></li>
					<li><a href="staff.asp?id=3">View/Edit Another Timetable</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	else
%>
	<div class="m_l">
		<div class="m_l_title">Staff Timetable</div>
		<div class="m_l_subtitle">View Or Edit <%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%>'s Timetable</div>
		<div class="m_l_ins">To Edit A Period, Just Click The Period You Wish To Change.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<%
			if var_est_enabled_weekends = "1" then
			%>
			<!--#include virtual="/pt/modules/ss/timetables/admin_7day.inc"-->
			<%
			else
			%>
			<!--#include virtual="/pt/modules/ss/timetables/admin_5day.inc"-->
			<%
			end if
			%>
			<div class="botopts">
				<ul>
					<li><a href="#" onmouseup="history.back();">Return</a></li>
					<li><a href="staff.asp?id=3">View/Edit Another Timetable</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	end if

	RSUSER.close
	set RSUSER = nothing

	else

	if session("sess_adminlevel") <> 1 then

	if session("sess_dept") <> "" then

	RSUSERSQL = "SELECT ID, LN, FN, TITLE, DEPT, CATEGORY, ENTITLEMENT FROM Timetables WHERE DEPT = " & session("sess_dept") & " ORDER BY LN"
		
	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<div class="m_l">
		<div class="m_l_title">Staff Timetables</div>
		<div class="m_l_subtitle">View And Edit People's Timetables.</div>
		<div class="m_l_ins">To Display A Timetable, Just Click The Person's You Wish To See.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
			<%
			do until RSUSER.EOF
			%>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="location.href='staff.asp?id=3&amp;user=<%=RSUSER("ID")%>'">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="<%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%>"></td>
					<td class="m_l_list_t"><b><%=RSUSER("LN")%>, <%=Left(RSUSER("FN"),1)%>.</b></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
				RSUSER.MOVENEXT
				loop
				%>
			</table>
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 10pt;">Displaying A Total Of <b><%=RSUSER.RECORDCOUNT%></b> Staff</div>
		</div>
	</div>
<%
				RSUSER.close
				set RSUSER = nothing
	else
%>
	<div class="m_l">
		<div class="m_l_title">Sorry!</div>
		<div class="m_l_subtitle">The System Can't Display Any Staff!</div>
		<div class="m_l_ins">The System Does Not Know What Your Department Is!</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Sorry, <%=session("sess_fn")%>, But No Staff Can Be Displayed Because Your Account Does Not Have A Department Linked To It.<br>
			<b>Please Contact The System Administrator.</b>
			<div class="botopts">
				<ul>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	end if

	else

	if (request("order")) = "LN" then
	RSUSERSQL = "SELECT ID, LN, FN, TITLE, DEPT, CATEGORY, ENTITLEMENT FROM Timetables ORDER BY LN"
	elseif (request("order")) = "FN" then
	RSUSERSQL = "SELECT ID, LN, FN, TITLE, DEPT, CATEGORY, ENTITLEMENT FROM Timetables ORDER BY FN"
	elseif (request("order")) = "DEPT" then
	RSUSERSQL = "SELECT Departments.SHORT, Timetables.ID, Timetables.TITLE, Timetables.LN, Timetables.FN, Timetables.DEPT, Timetables.DEFROOM, Timetables.ENTITLEMENT, Timetables.CATEGORY FROM Departments INNER JOIN Timetables ON Departments.DEPTID = Timetables.DEPT ORDER BY Departments.SHORT, Timetables.LN"
	elseif (request("order")) = "ENT" then
	RSUSERSQL = "SELECT ID, LN, FN, TITLE, DEPT, CATEGORY, ENTITLEMENT FROM Timetables ORDER BY ENTITLEMENT DESC"
	else
	RSUSERSQL = "SELECT ID, LN, FN, TITLE, DEPT, CATEGORY, ENTITLEMENT FROM Timetables ORDER BY LN"
	end if


		
	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<div class="m_l">
		<div class="m_l_title">Staff Timetables</div>
		<div class="m_l_subtitle">View And Edit People's Timetables.</div>
		<div class="m_l_ins">To Display A Timetable, Just Click The Person's You Wish To See.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<div style="width: 100%; text-align: center;">Order... <span class="m_l_order"><a href="staff.asp?id=3&amp;order=LN"><b>By Surname</b></a></span> :: <span class="m_l_order"><a href="staff.asp?id=3&amp;order=FN"><b>By First Name</b></a></span> :: <span class="m_l_order"><a href="staff.asp?id=3&amp;order=DEPT"><b>By Department</b></a></span> :: <span class="m_l_order"><a href="staff.asp?id=3&amp;order=ENT"><b>By Entitlement</b></a></span></div>
			<hr size="1">
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
			<%
			do until RSUSER.EOF
			%>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="location.href='staff.asp?id=3&amp;user=<%=RSUSER("ID")%>'">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="<%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%>"></td>
					<td class="m_l_list_t"><b><%=RSUSER("LN")%>, <%=Left(RSUSER("FN"),1)%>.</b></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
				RSUSER.MOVENEXT
				loop
				%>
			</table>
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 10pt;">Displaying A Total Of <b><%=RSUSER.RECORDCOUNT%></b> Staff</div>
		</div>
	</div>
<%
				RSUSER.close
				set RSUSER = nothing
	end if

	end if
%>
<%
elseif pagetype = "4" then
	if (request("print") = "1") then
		response.redirect "/pt/admin/print.asp?id=6"
	else
%>
	<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->
<%
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
			If Any Information On This List Is Incorrect, Please Select The Relevant Member Of Staff. You Will Then Be Taken To A Page Where You Can Change The Member's Entitlement.
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
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="location.href='staff.asp?id=5&amp;user=<%=RSUSER("ID")%>'">
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
	end if

elseif pagetype = "5" then
%>
	<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->
<%
	if (request("good")) = "1" then
%>
	<div class="m_l">
		<div class="m_l_title">Staff Details</div>
		<div class="m_l_subtitle">Congratulations, Details Updated!</div>
		<div class="m_l_ins">You Have Successfully Updated Staff Details!</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Congratulations, <%=session("sess_fn")%>! You Have Successfully Updated Staff Details.<br>
			These Changes Come In With Immediate Effect.
			<div class="botopts">
				<ul>
					<li><a href="staff.asp?id=5">Edit Someone Else's Details</a></li>
					<li><a href="staff.asp?id=6">Add A Member Of Staff</a></li>
					<li><a href="staff.asp?id=7">Delete A Member Of Staff</a></li>
					<li><a href="staff.asp?id=4">Goto Period Entitlements</a></li>
					<li><a href="staff.asp">Return To Staff Management</a></li>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	elseif (request("user")) <> "" then

	RSUSERSQL = "SELECT ID, LN, FN, TITLE, DEPT, CATEGORY, ENTITLEMENT, DEFROOM FROM Timetables WHERE ID = " & request("user")

	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
	
	if RSUSER.RECORDCOUNT = 0 THEN
%>
	<div class="m_l">
		<div class="m_l_title">Staff Details</div>
		<div class="m_l_subtitle">Cannot Find A User!</div>
		<div class="m_l_ins">The User ID Specified Is Invalid!</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Sorry, <%=session("sess_fn")%>, But We Cannot Find A User For ID <b>No.<%=request("user")%></b>.<br>
			Please Choose An Option From Below To Continue.
			<div class="botopts">
				<ul>
					<li><a href="#" onmouseup="history.back();">Return</a></li>
					<li><a href="staff.asp?id=5">Edit Another Member Of Staff's Details</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	else
%>
	<form name="edit" action="/pt/modules/ss/db/edit.asp?edittype=1" method="post">
	<div class="m_l">
		<div class="m_l_title">Staff Details</div>
		<div class="m_l_subtitle">Edit <%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%>'s Details</div>
		<div class="m_l_ins">Please Correct The Information Below, And Click "Save".</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<%
				do until RSUSER.EOF
				
				RSDEPTSQL = "SELECT * FROM Departments ORDER BY Short"
				RSROOMSQL = "SELECT * FROM Rooms ORDER BY ROOMNO ASC"
					
				Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
				RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic

				Set RSROOM = Server.CreateObject("Adodb.RecordSet")
				RSROOM.Open RSROOMSQL, dataconn, adopenkeyset, adlockoptimistic
				%>
				<tr class="m_l_list_b">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="<%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%>"></td>
					<td class="m_l_list_t"><b><%=RSUSER("LN")%>, <%=Left(RSUSER("FN"),1)%>.</b></td>
				</tr>
				<tr>
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Title</b></td>
								<td>
								<select class="b_std" size="1" name="TITLE">
									<option<%if RSUSER("TITLE") = "Mr" then%> selected<%else%><%end if%>>Mr</option>
									<option<%if RSUSER("TITLE") = "Mrs" then%> selected<%else%><%end if%>>Mrs</option>
									<option<%if RSUSER("TITLE") = "Miss" then%> selected<%else%><%end if%>>Miss</option>
									<option<%if RSUSER("TITLE") = "Ms" then%> selected<%else%><%end if%>>Ms</option>
									<option<%if RSUSER("TITLE") = "Mdme" then%> selected<%else%><%end if%>>Mdme</option>
									<option<%if RSUSER("TITLE") = "Dr" then%> selected<%else%><%end if%>>Dr</option>
								</select>								
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 150px; padding-left: 5px;"><b>First Name</b></td>
								<td><input class="b_std" type="text" name="FN" size="20" value="<%=RSUSER("FN")%>"><input type="hidden" name="USER" value="<%=RSUSER("ID")%>"></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Last Name</b></td>
								<td><input class="b_std" type="text" name="LN" size="20" value="<%=RSUSER("LN")%>"></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Department</b></td>
								<td>
								<select class="b_std" size="1" name="DEPT">
									<%
									do until RSDEPT.EOF
									%>
									<option value="<%=RSDEPT("DEPTID")%>"<%if RSDEPT("DEPTID") = RSUSER("DEPT") then%> selected<%else%><%end if%>><%=RSDEPT("Full")%></option>
									<%
									RSDEPT.MOVENEXT
									loop
									%>
								</select>
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Position</b></td>
								<td>
								<select class="b_std" size="1" name="CATEGORY">
									<option value="T"<%if RSUSER("CATEGORY") = "T" then%> selected<%else%><%end if%>>Teacher</option>
									<option value="PT"<%if RSUSER("CATEGORY") = "PT" then%> selected<%else%><%end if%>>Principal Teacher</option>
									<option value="DHT"<%if RSUSER("CATEGORY") = "DHT" then%> selected<%else%><%end if%>>Deputy Head Teacher</option>
									<option value="HT"<%if RSUSER("CATEGORY") = "HT" then%> selected<%else%><%end if%>>Head Teacher</option>
									<option value="PS"<%if RSUSER("CATEGORY") = "PS" then%> selected<%else%><%end if%>>Pupil Support</option>
									<option value="OC"<%if RSUSER("CATEGORY") = "OC" then%> selected<%else%><%end if%>>Outside Cover</option>
								</select>
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Usual Room</b></td>
								<td>
								<select class="b_std" size="1" name="DEFROOM">
									<option value="">N/A</option>
									<%
									do until RSROOM.EOF
									%>
									<option<%if RSROOM("ROOMNO") = RSUSER("DEFROOM") then%> selected<%else%><%end if%>><%=RSROOM("ROOMNO")%></option>
									<%
									RSROOM.MOVENEXT
									loop
									%>
								</select>
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Period Entitlement</b></td>
								<td style="font-size: 10pt;">
								<input class="b_std" type="text" name="ENTITLEMENT" size="2" maxlength="2" value="<%=RSUSER("ENTITLEMENT")%>" onblur="noalpha(this.value);"> Periods/Week
								</td>
							</tr>
							</table>
					<hr size="1">
					<%
					RSACCCHECKSQL = "SELECT * FROM Users WHERE TTID = " & RSUSER("ID")
					
					Set RSACCCHECK = Server.CreateObject("Adodb.RecordSet")
					RSACCCHECK.Open RSACCCHECKSQL, userconn, adopenkeyset, adlockoptimistic

					RSACCCHECK2SQL = "SELECT * FROM Admin WHERE TTID = " & RSUSER("ID")
					
					Set RSACCCHECK2 = Server.CreateObject("Adodb.RecordSet")
					RSACCCHECK2.Open RSACCCHECK2SQL, userconn, adopenkeyset, adlockoptimistic
					%>
					<div style="width: 100%; text-align: center;"><b><a href="#" onmouseup="document.edit.submit();">Save</a></b> :: <b><a href="#" onmouseup="document.edit.reset();">Reset</a></b> :: <%if RSACCCHECK.RECORDCOUNT => 1 then%><a href="staff.asp?id=8&amp;uid=<%=RSUSER("ID")%>&amp;manage=1"><b>Manage Account</b></a><%elseif RSACCCHECK2.RECORDCOUNT => 1 then%><a href="staff.asp?id=8&amp;uid=<%=RSUSER("ID")%>&amp;manage=1"><b>Manage Account</b></a><%else%><a href="staff.asp?id=8&amp;uid=<%=RSUSER("ID")%>"><b>Create Account</b></a><%end if%> :: <b><a href="staff.asp?id=3&amp;user=<%=RSUSER("ID")%>">View/Edit Timetable</a></b> :: <b><a href="staff.asp?id=7&amp;confirm=1&amp;user=<%=RSUSER("ID")%>">Delete This Person</a></b></div>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="3"></td>
				</tr>
				<%
					RSACCCHECK.close
					set RSACCCHECK = nothing
					RSACCCHECK2.close
					set RSACCCHECK2 = nothing

				RSDEPT.close
				RSROOM.close
				set RSDEPT = nothing
				set RSROOM = nothing
				
				RSUSER.MOVENEXT
				loop
				%>
			</table>
		</div>
	</div>
	</form>
<%
	end if

	RSUSER.close
	set RSUSER = nothing

	else

	RSUSERSQL = "SELECT ID, LN, FN, TITLE, DEPT, CATEGORY, ENTITLEMENT FROM Timetables ORDER BY LN"
		
	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<div class="m_l">
		<div class="m_l_title">Staff Details</div>
		<div class="m_l_subtitle">Edit Staff Details, Such As Their Names.</div>
		<div class="m_l_ins">To Edit Someone's Details, Simply Click Their Name.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
			<%
			do until RSUSER.EOF
			%>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="location.href='staff.asp?id=5&amp;user=<%=RSUSER("ID")%>'">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="<%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%>"></td>
					<td class="m_l_list_t"><b><%=RSUSER("LN")%>, <%=Left(RSUSER("FN"),1)%>.</b></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
				RSUSER.MOVENEXT
				loop
				%>
			</table>
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 10pt;">Displaying A Total Of <b><%=RSUSER.RECORDCOUNT%></b> Staff</div>
		</div>
	</div>

<%
	RSUSER.close
	set RSUSER = nothing

	end if

elseif pagetype = "6" then
%>
	<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->
<%
	if (request("err")) = "1" then
%>
	<div class="m_l">
		<div class="m_l_title">Add A Member Of Staff</div>
		<div class="m_l_subtitle">Cannnot Add This Member!</div>
		<div class="m_l_ins">The Person Already Exists!</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Sorry, <%=session("sess_fn")%>, But We Cannot Add This Person To The Database Because They Already Exist In It.<br>
			Please Choose An Option From Below To Continue.
			<div class="botopts">
				<ul>
					<li><a href="#" onmouseup="history.back();">Go Back And Change Details</a></li>
					<li><a href="staff.asp">Return To The Staff Management Home</a></li>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	elseif (request("err")) = "2" then
%>
	<div class="m_l">
		<div class="m_l_title">Add A Member Of Staff</div>
		<div class="m_l_subtitle">Cannnot Add This Member!</div>
		<div class="m_l_ins">You Left Some Information Out!</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Sorry, <%=session("sess_fn")%>, But We Cannot Add This Person To The Database As Some Information Is Missing.<br>
			All Fields Are Required For A Successful Addition.
			<div class="botopts">
				<ul>
					<li><a href="#" onmouseup="history.back();">Go Back And Amend Details</a></li>
					<li><a href="staff.asp">Return To The Staff Management Home</a></li>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	elseif (request("good")) = "1" then
%>
	<div class="m_l">
		<div class="m_l_title">Add A Member Of Staff</div>
		<div class="m_l_subtitle">Congratulations, Person Added!</div>
		<div class="m_l_ins">You Have Successfully Added This Person To The Database!</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->		
			Congratulations, <%=session("sess_fn")%>! You Have Successfully Added These Staff Details To The Database.<br>
			The New Member Of Staff Can Now Be Used For Please Takes.<br>
			<b>However, Before You Do That, It Is Highly Recommended That You Add Their Timetable Information In.</b>
			<br><span style="font-size: 14pt; font-weight: bold;"><a href="staff.asp?id=3&amp;user=<%=request("user")%>">Click Here To Add The New Member's Timetable Information</a></span>
			<div class="botopts">
				<ul>
					<li><a href="staff.asp?id=6">Add Another Member Of Staff</a></li>
					<li><a href="staff.asp">Return To Staff Management</a></li>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	else
%>
	<form name="add" action="/pt/modules/ss/db/add.asp?addtype=1" method="post">
	<div class="m_l">
		<div class="m_l_title">Add A Member Of Staff</div>
		<div class="m_l_subtitle">Add A Member Of Staff To The System.</div>
		<div class="m_l_ins">Please Fill Out The Information Below, And Click "Add".</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<%
			RSDEPTSQL = "SELECT * FROM Departments ORDER BY SHORT"
			RSROOMSQL = "SELECT * FROM Rooms ORDER BY ROOMNO ASC"
				
			Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
			RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic

			Set RSROOM = Server.CreateObject("Adodb.RecordSet")
			RSROOM.Open RSROOMSQL, dataconn, adopenkeyset, adlockoptimistic
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="New Member Of Staff"></td>
					<td class="m_l_list_t"><b>New Member Of Staff</b></td>
				</tr>
				<tr>
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Title</b></td>
								<td>
								<select class="b_std" size="1" name="TITLE">
									<option>Mr</option>
									<option>Mrs</option>
									<option>Miss</option>
									<option>Ms</option>
									<option>Mdme</option>
									<option>Dr</option>
								</select>								
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 150px; padding-left: 5px;"><b>First Name</b></td>
								<td><input class="b_std" type="text" name="FN" size="20"></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Last Name</b></td>
								<td><input class="b_std" type="text" name="LN" size="20"></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Department</b></td>
								<td>
								<select class="b_std" size="1" name="DEPT">
									<%
									do until RSDEPT.EOF
									%>
									<option value="<%=RSDEPT("DEPTID")%>"><%=RSDEPT("Full")%></option>
									<%
									RSDEPT.MOVENEXT
									loop
									%>
								</select>
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Position</b></td>
								<td>
								<select class="b_std" size="1" name="CATEGORY">
									<option value="T">Teacher</option>
									<option value="PT">Principal Teacher</option>
									<option value="DHT">Deputy Head Teacher</option>
									<option value="HT">Head Teacher</option>
									<option value="PS">Pupil Support</option>
									<option value="OC">Outside Cover</option>
								</select>
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Usual Room</b></td>
								<td>
								<select class="b_std" size="1" name="DEFROOM">
									<option value="">N/A</option>
									<%
									do until RSROOM.EOF
									%>
									<option><%=RSROOM("ROOMNO")%></option>
									<%
									RSROOM.MOVENEXT
									loop
									%>
								</select>
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Period Entitlement</b></td>
								<td style="font-size: 10pt;">
								<input class="b_std" type="text" name="ENTITLEMENT" size="2" maxlength="2" onblur="noalpha(this.value);"> Periods/Week
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Create Account?</b></td>
								<td style="font-size: 10pt;">
								<div style="position: relative; right: 4px;"><input type="checkbox" value="1" name="CREATE" id="create" onmouseup="showdetail('em'); showdetail('em_sep'); showdetail('un'); showdetail('un_sep'); showdetail('pw_sep'); showdetail('pw'); showdetail('type_sep'); showdetail('type'); <%if var_est_enabled_pin = "1" then%>showdetail('pin_sep'); showdetail('pin');<%else%><%end if%>">&nbsp;Check The Box For Yes - Uncheck For No</div>
								</td>
							</tr>
							<tr id="list_un_sep" style="display: none;">
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr id="list_un" style="display: none;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Username</b></td>
								<td style="font-size: 10pt;">
								<input class="b_std" type="text" name="UN" size="20">
								</td>
							</tr>
							<tr id="list_pw_sep" style="display: none;">
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr id="list_pw" style="display: none;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Password</b></td>
								<td style="font-size: 10pt;">
								<input class="b_pwd" type="password" name="PW" size="20">
								</td>
							</tr>
							<%
							if var_est_enabled_pin = "1" then
							%>
							<tr id="list_pin_sep" style="display: none;">
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr id="list_pin" style="display: none;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>PIN</b></td>
								<td style="font-size: 10pt;">
								<input class="b_pwd" type="password" name="p1" size="3" maxlength="3" onkeyup="pjump();">&nbsp;/&nbsp;<input class="b_pwd" type="password" name="p2" size="3" maxlength="3">
								</td>
							</tr>
							<%
							else
							end if
							%>
							<tr id="list_em_sep" style="display: none;">
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr id="list_em" style="display: none;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>E-Mail</b></td>
								<td style="font-size: 10pt;">
								<input class="b_std" type="text" name="email" size="20">&nbsp;@<%=var_emaildomain1%>
								</td>
							</tr>
							<tr id="list_type_sep" style="display: none;">
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr id="list_type" style="display: none;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Account Type</b></td>
								<td style="position: relative; right: 4px; font-size: 10pt;">
								<div style="padding: 2px 0px 2px 0px;"><input type="radio" value="1" name="TYPE" id="create" checked>&nbsp;Standard <%=var_usernames_std%> Account</div>
								<div style="padding: 2px 0px 2px 0px;"><input type="radio" value="2" name="TYPE" id="create">&nbsp;Restricted <%=var_usernames_admin%> Account</div>
								<div style="padding: 2px 0px 2px 0px;"><input type="radio" value="3" name="TYPE" id="create">&nbsp;Full-Access <%=var_usernames_admin%> Account</div>
								</td>
							</tr>
						</table>
					<hr size="1">
					<div style="width: 100%; text-align: center;"><b><a href="#" onmouseup="document.add.submit();">Add</a></b> :: <b><a href="#" onmouseup="document.add.reset();">Clear</a></b></div>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="3"></td>
				</tr>
				<%
				RSDEPT.close
				RSROOM.close
				set RSDEPT = nothing
				set RSROOM = nothing
				%>
			</table>
		</div>
	</div>
	</form>
<%
	end if
	
elseif pagetype = "7" then
%>
	<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->
<%
	if (request("confirm")) = "1" then

	RSUSERSQL = "SELECT ID, LN, FN, TITLE FROM Timetables WHERE ID = " & (request("user"))

	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<div class="m_l">
		<div class="m_l_title">Delete A Member Of Staff</div>
		<div class="m_l_subtitle">Are You Sure?</div>
		<div class="m_l_ins">Please Check This Is The Person You Wish To Remove.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->		
			Please Ensure That The Member Of Staff You Wish To Delete Is:<br>
			<span style="font-size: 14pt; font-weight: bold;"><%=RSUSER("TITLE")%>. <%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%></span><br>
			Is This Correct? Clicking "No" Returns You To The Previous Page, Whilst "Yes" Deletes <%=RSUSER("FN")%>'s Entry...
			<p>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td style="font-size: 22pt; font-weight: bold; text-align: center;" onmouseup="location.href='/pt/modules/ss/db/delete.asp?deltype=1&amp;user=<%=RSUSER("ID")%>'">Yes</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="1"></td>
				</tr>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td style="font-size: 22pt; font-weight: bold; text-align: center;" onmouseup="history.back();">No</td>
				</tr>
			</table>
			<p>
			...Or, Choose One Of These Options Below.
			<div class="botopts">
				<ul>
					<li><a href="staff.asp">Return To Staff Management</a></li>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	RSUSER.close
	set RSUSER = nothing

	elseif (request("good")) = "1" then
%>
	<div class="m_l">
		<div class="m_l_title">Delete A Member Of Staff</div>
		<div class="m_l_subtitle">Congratulations, Person Removed!</div>
		<div class="m_l_ins">You Have Successfully Deleted The Member From The System's Database!</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Congratulations, <%=session("sess_fn")%>! You Have Successfully Removed The Member Of Staff From The Database.<br>
			<div class="botopts">
				<ul>
					<li><a href="staff.asp?id=7">Delete Another Member Of Staff</a></li>
					<li><a href="staff.asp">Return To Staff Management</a></li>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	else

	RSUSERSQL = "SELECT ID, LN, FN, TITLE, DEPT, CATEGORY, ENTITLEMENT, DEFROOM FROM Timetables ORDER BY LN ASC"

	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<div class="m_l">
		<div class="m_l_title">Delete A Member Of Staff</div>
		<div class="m_l_subtitle">Permanatley Remove A Person From The Database.</div>
		<div class="m_l_ins">Please Select The Person You Wish To Delete.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
			<%
			do until RSUSER.EOF
			%>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="location.href='staff.asp?id=7&amp;confirm=1&amp;user=<%=RSUSER("ID")%>'">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="<%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%>"></td>
					<td class="m_l_list_t"><b><%=RSUSER("LN")%>, <%=Left(RSUSER("FN"),1)%>.</b></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
				RSUSER.MOVENEXT
				loop
				%>
			</table>
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 10pt;">Displaying A Total Of <b><%=RSUSER.RECORDCOUNT%></b> Staff</div>
		</div>
	</div>
<%
	RSUSER.close
	set RSUSER = nothing

	end if

elseif pagetype = "8" then

	if (request("gd")) = "1" then
%>
	<div class="m_l">
		<div class="m_l_title">Account Created!</div>
		<div class="m_l_subtitle">Congratulations!</div>
		<div class="m_l_ins">You Have Successfully Created A User Account!</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Congratulations, <%=session("sess_fn")%>! You Have Successfully Created A User Account.<br>
			The User Can Now Login Using The Relevant Login Screen And Details You Have Just Entered.
			<div class="botopts">
				<ul>
					<li><a href="staff.asp?id=2">Return To The Staff Information List</a></li>
					<li><a href="staff.asp?id=1">Return To The Staff Management Homepage</a></li>
					<li><a href="default.asp?id=1">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	elseif (request("uid")) = "" then
		response.redirect "/pt/admin/staff.asp?id=2"

	elseif (request("manage")) = "1" then

		RSUSERSQL = "SELECT ID, FN, LN FROM Timetables WHERE id = " & request("uid")
	
		Set RSUSER = Server.CreateObject("Adodb.RecordSet")
		RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
		
		if (request("del")) = "1" then
%>
	<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->
	<div class="m_l">
		<div class="m_l_title">Delete Account</div>
		<div class="m_l_subtitle">Are You Sure?</div>
		<div class="m_l_ins">Please Select From Below.</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Please Note That This Will Only <b>Delete The <i>Access</i> Account From The System, Meaning <%=RSUSER("FN")%> Will No Longer Be Able To Login.</b><br>
			<%=RSUSER("FN")%>'s Timetable Details Will Still Remain Online. To Remove Everything, Click <a href="staff.asp?id=7&amp;confirm=1&amp;user=<%=RSUSER("ID")%>"><b>Here</b></a>.<br>
			Do You Want To Proceed?
			<p>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td style="font-size: 22pt; font-weight: bold; text-align: center;" onmouseup="location.href='/pt/modules/ss/db/delete.asp?deltype=4&amp;uid=<%=RSUSER("ID")%>'">Yes</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="1"></td>
				</tr>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td style="font-size: 22pt; font-weight: bold; text-align: center;" onmouseup="history.back();">No</td>
				</tr>
			</table>
			<p>
			...Or, Choose One Of These Options Below.
			<div class="botopts">
				<ul>
					<li><a href="staff.asp">Return To Staff Management</a></li>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
		elseif (request("del")) = "2" then
%>
	<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->
	<div class="m_l">
		<div class="m_l_title">Account Deleted!</div>
		<div class="m_l_subtitle">The Account You Wanted Gone Is No More!</div>
		<div class="m_l_ins">The Member's Account You Removed Can No Longer Login.</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Since You Have Just Removed The Staff Member's Access Account, The Member Is Now Unable To Login Using That Account.<br>
			The Account Has Been Successfully Removed From The System.
			<div class="botopts">
				<ul>
					<li><a href="staff.asp">Return To Staff Management</a></li>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
		else

		RSDETAILSSQL = "SELECT * FROM Users WHERE TTID = " & request("UID")
	
		Set RSDETAILS = Server.CreateObject("Adodb.RecordSet")
		RSDETAILS.Open RSDETAILSSQL, userconn, adopenkeyset, adlockoptimistic
		
			if RSDETAILS.RECORDCOUNT = 0 then

				RSDETAILS.close
				set RSDETAILS = nothing

				RSDETAILSSQL = "SELECT * FROM Admin WHERE TTID = " & request("UID")
	
				Set RSDETAILS = Server.CreateObject("Adodb.RecordSet")
				RSDETAILS.Open RSDETAILSSQL, userconn, adopenkeyset, adlockoptimistic
				
				table = "2"
				acclevel = RSDETAILS("ACCLEVEL")
			
			else
				table = "1"
				acclevel = ""
			end if
%>
	<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->
	<form name="add" action="/pt/modules/ss/db/edit.asp?edittype=9&amp;uid=<%=RSUSER("ID")%>" method="post">
	<div class="m_l">
		<div class="m_l_title">Manage Account</div>
		<div class="m_l_subtitle">Manage <%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%>'s Account</div>
		<div class="m_l_ins">You Can Change Account Details, Or Simply Delete The Account.</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<%
			if (request("err")) = "1" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Missed One Or More Fields! <b>All</b> Fields Must Be Entered!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "2" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">The Username You Used Has Already Been Taken! Please Try Another!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "3" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">A Critical Error Has Occured! Please Contact The System Administrator!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			else
			end if
			%>

			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="location.href='staff.asp?id=8&amp;uid=<%=RSUSER("ID")%>&amp;manage=1&amp;del=1'">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_del.gif" border="0" alt="Delete This Member's Account"></td>
					<td class="m_l_list_t"><b>Delete The User Account For <%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%></b></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('details');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="Editing This Member Of Staff's Details"></td>
					<td class="m_l_list_t"><b>Edit Account Details</b></td>
				</tr>
				<tr id="list_details">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Username</b></td>
								<td style="font-size: 10pt;">
								<input class="b_std" type="text" name="UN" size="20" value="<%=RSDETAILS("UN")%>">
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Password</b></td>
								<td style="font-size: 10pt;">
								<input class="b_pwd" type="password" name="PW" size="20">&nbsp;Not Entered For Security Reasons
								</td>
							</tr>
							<%
							if var_est_enabled_pin = "1" then
							%>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>PIN</b></td>
								<td style="font-size: 10pt;">
								<input class="b_pwd" type="password" name="p1" size="3" maxlength="3" onkeyup="pjump();">&nbsp;/&nbsp;<input class="b_pwd" type="password" name="p2" size="3" maxlength="3">&nbsp;Not Entered For Secuity Reasons
								</td>
							</tr>
							<%
							else
							end if
							%>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>E-Mail</b></td>
								<td style="font-size: 10pt;">
								<input class="b_std" type="text" name="email" size="20" value="<%=removeDomain(RSDETAILS("email"))%>">&nbsp;@<%=var_emaildomain1%>
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Account Type</b></td>
								<td style="position: relative; right: 4px; font-size: 10pt;">
								<div style="padding: 2px 0px 2px 0px;"><input type="radio" value="1" name="TYPE"<%if table = "1" then%> checked<%else%><%end if%>>&nbsp;Standard <%=var_usernames_std%> Account</div>
								<div style="padding: 2px 0px 2px 0px;"><input type="radio" value="2" name="TYPE"<%if (table = "2") AND (acclevel) = "2" then%> checked<%else%><%end if%>>&nbsp;Restricted <%=var_usernames_admin%> Account</div>
								<div style="padding: 2px 0px 2px 0px;"><input type="radio" value="3" name="TYPE"<%if (table = "2") AND (acclevel) = "1" then%> checked<%else%><%end if%>>&nbsp;Full-Access <%=var_usernames_admin%> Account</div>
								</td>
							</tr>
						</table>
					<hr size="1">
					<div style="width: 100%; text-align: center;"><b><a href="#" onmouseup="document.add.submit();">Save New Details</a></b> :: <b><a href="#" onmouseup="document.add.reset();">Reset</a></b></div>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="3"></td>
				</tr>
			</table>

		</div>
	</div>
	</form>
<%
		RSDETAILS.close
		set RSDETAILS = nothing

		end if

		RSUSER.close
		set RSUSER = nothing
	else

		RSUSERSQL = "SELECT ID, FN, LN FROM Timetables WHERE id = " & request("uid")
	
		Set RSUSER = Server.CreateObject("Adodb.RecordSet")
		RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
		
		if RSUSER.RECORDCOUNT = 0 then
			response.redirect "/pt/admin/staff.asp?id=2"
		else
		end if
%>
	<form name="add" action="/pt/modules/ss/db/add.asp?addtype=6&amp;uid=<%=RSUSER("ID")%>" method="post">
	<div class="m_l">
		<div class="m_l_title">Create An Account</div>
		<div class="m_l_subtitle">Create A User Account For <%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%>.</div>
		<div class="m_l_ins">Please Fill Out The Information Below, And Click "Add".</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<%
			if (request("err")) = "1" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Missed One Or More Fields! <b>All</b> Fields Must Be Entered!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "2" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">The Username You Used Has Already Been Taken! Please Try Another!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "3" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">A Critical Error Has Occured! Please Contact The System Administrator!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			else
			end if
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="New Member Of Staff"></td>
					<td class="m_l_list_t"><b>New Account Details For <%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%></b></td>
				</tr>
				<tr>
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Username</b></td>
								<td style="font-size: 10pt;">
								<input class="b_std" type="text" name="UN" size="20" value="<%=RSUSER("LN")%><%=left(RSUSER("FN"),1)%>">&nbsp;Initial Value Only A Suggestion!
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Password</b></td>
								<td style="font-size: 10pt;">
								<input class="b_pwd" type="password" name="PW" size="20">
								</td>
							</tr>
							<%
							if var_est_enabled_pin = "1" then
							%>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>PIN</b></td>
								<td style="font-size: 10pt;">
								<input class="b_pwd" type="password" name="p1" size="3" maxlength="3" onkeyup="pjump();">&nbsp;/&nbsp;<input class="b_pwd" type="password" name="p2" size="3" maxlength="3">
								</td>
							</tr>
							<%
							else
							end if
							%>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>E-Mail</b></td>
								<td style="font-size: 10pt;">
								<input class="b_std" type="text" name="email" size="20">&nbsp;@<%=var_emaildomain1%>
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Account Type</b></td>
								<td style="position: relative; right: 4px; font-size: 10pt;">
								<div style="padding: 2px 0px 2px 0px;"><input type="radio" value="1" name="TYPE" id="create" checked>&nbsp;Standard <%=var_usernames_std%> Account</div>
								<div style="padding: 2px 0px 2px 0px;"><input type="radio" value="2" name="TYPE" id="create">&nbsp;Restricted <%=var_usernames_admin%> Account</div>
								<div style="padding: 2px 0px 2px 0px;"><input type="radio" value="3" name="TYPE" id="create">&nbsp;Full-Access <%=var_usernames_admin%> Account</div>
								</td>
							</tr>
						</table>
					<hr size="1">
					<div style="width: 100%; text-align: center;"><b><a href="#" onmouseup="document.add.submit();">Add</a></b> :: <b><a href="#" onmouseup="document.add.reset();">Clear</a></b></div>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="3"></td>
				</tr>
			</table>
		</div>
	</div>
	</form>
<%
		RSUSER.close
		set RSUSER = nothing
	end if
else
%>
	<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->
	<div class="m_l">
		<div class="m_l_title">Staff Management</div>
		<div class="m_l_subtitle">Add, Delete Or Edit Staff Details.</div>
		<div class="m_l_ins">Please Choose A Task From Below...</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_t_l"><a href="staff.asp?id=2"><img src="/pt/media/icons/48_book.png" border="0" alt="Quickly View Staff Information By Using This Tool."></a></td>
					<td class="m_l_sel_t_r"><a href="staff.asp?id=2">View Staff Information List</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="staff.asp?id=3"><img src="/pt/media/icons/48_cal.png" border="0" alt="View And/Or Edit Staff Timetables Quickly And Easily."></a></td>
					<td class="m_l_sel_t_r"><a href="staff.asp?id=3">View/Edit Staff Timetables</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="staff.asp?id=4"><img src="/pt/media/icons/48_ent.png" border="0" alt="View And/Or Edit How Many Periods Staff Get Per Week For Please Takes."></a></td>
					<td class="m_l_sel_t_r"><a href="staff.asp?id=4">View/Edit Period Entitlements</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="staff.asp?id=6"><img src="/pt/media/icons/48_man.png" border="0" alt="Add A Member Of Staff To The Please Takes System By Filling Out A Quick Form."></a></td>
					<td class="m_l_sel_t_r"><a href="staff.asp?id=6">Add A Member Of Staff</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="staff.asp?id=7"><img src="/pt/media/icons/48_del.png" border="0" alt="When A Member Of Staff Leaves, Ensure That Their Record Is Removed, By Using This Tool."></a></td>
					<td class="m_l_sel_t_r"><a href="staff.asp?id=7">Delete A Member Of Staff</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="staff.asp?id=5"><img src="/pt/media/icons/48_edit.png" border="0" alt="Using This Tool, You May Edit Any Incorrect Information Held About Staff."></a></td>
					<td class="m_l_sel_t_r"><a href="staff.asp?id=5">Edit Staff Details</a></td>
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