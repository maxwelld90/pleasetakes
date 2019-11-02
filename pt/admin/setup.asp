<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" >
<!--#include virtual="/pt/modules/ss/usersys/logincheck_admin.inc"-->

<!--#include virtual="/pt/modules/ss/p_s.inc"-->
<%
if session("sess_un") <> settingsXML.documentElement.childNodes.item(0).childNodes.item(4).getAttribute("firstloginacc") then
	response.redirect "default.asp?err=2"
else

pagetype = request("id")
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="../modules/css/admin.css">
<script language="javascript" type="text/javascript" src="/pt/modules/js/admin.js"></script>
<%if (pagetype = "10") then%>
<script language="javascript" type="text/javascript">
function compredir()
{
 location.href = '/pt/modules/ss/db/setup.asp?setuptype=10';
}
</script>
<%
else
end if
%>
<title><%=var_ptitle%></title>
</head>

<body<%if (pagetype = "2") then%> onload="document.add.FULL.focus();"<%elseif (pagetype = "3") then%> onload="document.add.ROOM.focus();"<%elseif (pagetype = "10") then%> onLoad="setTimeout('compredir()', 7000)"<%else%><%end if%>>

<div class="smlb_b"></div>
<div class="topb_b"></div>

<%
if pagetype = "2" then

	RSCHECKSQL = "SELECT * FROM Departments"
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic

	if RSCHECK.RECORDCOUNT = 0 then
		RSCHECK.close
		set RSCHECK = nothing
		session("whereiam") = 2
	else
		RSCHECK.close
		set RSCHECK = nothing
		session("whereiam") = 3
	end if
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li><a href="setup.asp?id=1">Welcome</a> :: </li>
			<li><a href="setup.asp?id=2"><b>Step 1</b></a> :: </li>
			<%
			RSDEPTSQL = "SELECT * FROM Departments"
			Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
			RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic

			if RSDEPT.RECORDCOUNT => 1 then
			%>
			<li><a href="setup.asp?id=3">Step 2</a> :: </li>
			<%
			else
			%>
			<li>Step 2 :: </li>
			<%
			end if
			%>
			<li>Step 3 :: </li>
			<li>Step 4 :: </li>
			<li>Step 5 :: </li>
			<li>Step 6 :: </li>
			<li>Step 7 :: </li>
			<li>Complete</li>
			<%
			RSDEPT.close
			set RSDEPT = nothing
			%>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<form name="add" action="/pt/modules/ss/db/setup.asp?setuptype=1" method="post">
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Step 1</div>
		<div class="m_l_subtitle">Sorting Out Departments</div>
		<div class="m_l_ins">Please Read The Instructions Below And Once You Are Done, Click "Next".</div>
		<div class="m_l_sel" style="width: 100%;">
			<%
			if (request("err")) = "1" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Must Specify A Name For The Department You Wish To Add!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "2" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Must Add At Least <b>One</b> Department Before Continuing!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			else
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">Any Problems With Your Information Will Be Displayed Here.</td>
				</tr>
			</table>
			<hr size="1">
			<%
			end if
			%>
		The First Step Requires You To Add All The Departments <%=var_est_full%> Has.<br>
		The Form Below Allows You To Add Departments, And Once You Click "Add", They Will Appear Underneath In Alphabetical Order.<br>
		Once You Have Added All The Departments, Click "Next".<br>

			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('add');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_dept.gif" border="0" alt="Add A Department"></td>
					<td class="m_l_list_t"><b>Add A Department</b></td>
				</tr>
				<tr id="list_add">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">

						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 140px; padding-left: 5px;"><b>Full Name</b></td>
								<td colspan="2" style="width: 140px;"><input class="b_std" type="text" name="FULL" size="20"></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 140px; padding-left: 5px;"><b>Short Name</b></td>
								<td class="m_l_list_t" style="width: 140px;"><input class="b_std" type="text" name="SHORT" size="20"></td>
								<td class="m_l_list_t">If Left Blank, The Department Will Have No Short Name.</td>
							</tr>
							<tr>
								<td colspan="3" class="m_l_list_t">
								<hr size="1">
								<div style="width: 100%; text-align: center;"><b><a href="#" onmouseup="document.add.submit();">Save</a></b> :: <b><a href="#" onmouseup="document.add.reset();">Clear</a></b></div>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep"></td>
				</tr>
			</table>

			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('added');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_dept.gif" border="0" alt="Successfully Added Departments"></td>
					<td class="m_l_list_t"><b>Successfully Added Departments</b></td>
				</tr>
				<tr id="list_added" style="display: none;">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">

						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
						<%
						RSDEPTSQL = "SELECT * FROM Departments ORDER BY SHORT"
						Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
						RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic

						if RSDEPT.RECORDCOUNT => 1 then

						do until RSDEPT.EOF
						%>
							<tr class="m_l_list_b" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);">
								<td class="m_l_list_t" style="height: 26px; padding-left: 5px;"><b><%=RSDEPT("FULL")%></b><%if RSDEPT("SHORT") = RSDEPT("FULL") then%><%else%> (<%=RSDEPT("SHORT")%>)<%end if%></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						<%
						RSDEPT.MOVENEXT
						loop

						else
						%>
							<tr class="m_l_list_b" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);">
								<td class="m_l_list_t" style="height: 26px; padding-left: 5px;">No Departments Found!</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						<%
						end if

						RSDEPT.close
						set RSDEPT = nothing
						%>
							</table>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep"></td>
				</tr>
			</table>
		<hr size="1">
		<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="setup.asp?id=3">Next></a></div>
		</div>
	</div>
	</form>
<%
elseif pagetype = "3" then

	RSCHECKSQL = "SELECT * FROM ROOMS"
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic

	if RSCHECK.RECORDCOUNT = 0 then
		RSCHECK.close
		set RSCHECK = nothing
		session("whereiam") = 3
	else
		RSCHECK.close
		set RSCHECK = nothing
		session("whereiam") = 4
	end if

	if session("whereiam") < 3 then
		response.redirect "setup.asp?err=1"
	else
	end if

	RSCHECKSQL = "SELECT * FROM Departments"
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic

	if RSCHECK.RECORDCOUNT = 0 then
		RSCHECK.close
		set RSCHECK = nothing
		response.redirect "setup.asp?id=2&err=2"
	else
		RSCHECK.close
		set RSCHECK = nothing
	end if
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li><a href="setup.asp?id=1">Welcome</a> :: </li>
			<li><a href="setup.asp?id=2">Step 1</a> :: </li>
			<li><a href="setup.asp?id=3"><b>Step 2</b></a> :: </li>
			<%
			RSROOMSQL = "SELECT * FROM ROOMS ORDER BY ROOMNO"
			Set RSROOM = Server.CreateObject("Adodb.RecordSet")
			RSROOM.Open RSROOMSQL, dataconn, adopenkeyset, adlockoptimistic

			if RSROOM.RECORDCOUNT => 1 then
			%>
			<li><a href="setup.asp?id=4">Step 3</a> :: </li>
			<%
			else
			%>
			<li>Step 3 :: </li>
			<%
			end if
			%>
			<li>Step 4 :: </li>
			<li>Step 5 :: </li>
			<li>Step 6 :: </li>
			<li>Step 7 :: </li>
			<li>Complete</li>
			<%
			RSROOM.close
			set RSROOM = nothing
			%>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<form name="add" action="/pt/modules/ss/db/setup.asp?setuptype=2" method="post">
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Step 2</div>
		<div class="m_l_subtitle">Adding Rooms</div>
		<div class="m_l_ins">Please Add Rooms From Below, And Click "Next" When You Are Done.</div>
		<div class="m_l_sel" style="width: 100%;">
			<%
			if (request("err")) = "1" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Must Specify A Name Or Number For The Room You Wish To Add!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "2" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Must Add At Least <b>One</b> Room Before Continuing!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "3" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">The Room You Specified Already Exists!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			else
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">Any Problems With Your Information Will Be Displayed Here.</td>
				</tr>
			</table>
			<hr size="1">
			<%
			end if
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('add');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_room.gif" border="0" alt="Add A Room"></td>
					<td class="m_l_list_t"><b>Add A Room</b></td>
				</tr>
				<tr id="list_add">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">

						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 140px; padding-left: 5px;"><b>Room Name/Number</b></td>
								<td ><input class="b_std" type="text" name="ROOM" size="20"></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2" class="m_l_list_t">
								<hr size="1">
								<div style="width: 100%; text-align: center;"><b><a href="#" onmouseup="document.add.submit();">Save</a></b> :: <b><a href="#" onmouseup="document.add.reset();">Clear</a></b></div>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep"></td>
				</tr>
			</table>

			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('added');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_room.gif" border="0" alt="Successfully Added Rooms"></td>
					<td class="m_l_list_t"><b>Successfully Added Rooms</b></td>
				</tr>
				<tr id="list_added" style="display: none;">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">

						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
						<%
						RSROOMSQL = "SELECT * FROM Rooms ORDER BY ROOMNO ASC"
						Set RSROOM = Server.CreateObject("Adodb.RecordSet")
						RSROOM.Open RSROOMSQL, dataconn, adopenkeyset, adlockoptimistic

						if RSROOM.RECORDCOUNT = 0 then
						%>
							<tr class="m_l_list_b" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);">
								<td class="m_l_list_t" style="height: 26px; padding-left: 5px;">No Rooms Found!</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						<%
						else

						do until RSROOM.EOF
						%>
							<tr class="m_l_list_b" >
								<td class="m_l_list_t" style="width: 200px; height: 26px; padding-left: 5px;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);"><%=RSROOM("ROOMNO")%></td>
						<%
						if RSROOM.EOF then
						else
						RSROOM.MOVENEXT
						end if
						%>
								<td style="width: 25px; background-color: #FFF;"></td>
						<%
						if RSROOM.EOF then
						else
						%>
								<td class="m_l_list_t" style="width: 232px; height: 26px; padding-left: 5px;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);"><%=RSROOM("ROOMNO")%></td>
						<%
						end if
						if RSROOM.EOF then
						else
						RSROOM.MOVENEXT
						end if
						%>
								<td style="width: 25px; background-color: #FFF;"></td>
						<%
						if RSROOM.EOF then
						else
						%>
								<td class="m_l_list_t" style="width: 232px; height: 26px; padding-left: 5px;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);"><%=RSROOM("ROOMNO")%></td>
						<%
						end if
						%>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						<%
						if RSROOM.EOF then
						else
						RSROOM.MOVENEXT
						end if

						loop
						end if

						RSROOM.close
						set ROOM = nothing
						%>
							</table>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep"></td>
				</tr>
			</table>
		<hr size="1">
		<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="setup.asp?id=4">Next></a></div>
		</div>
	</div>
	</form>
<%
elseif pagetype = "4" then

	RSCHECKSQL = "SELECT * FROM Timetables"
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic

	if RSCHECK.RECORDCOUNT = 0 then
		RSCHECK.close
		set RSCHECK = nothing
		session("whereiam") = 4
	else
		RSCHECK.close
		set RSCHECK = nothing
		session("whereiam") = 5
	end if

	if session("whereiam") < 4 then
		response.redirect "setup.asp?err=1"
	else
	end if

	RSCHECKSQL = "SELECT * FROM Rooms"
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic

	if RSCHECK.RECORDCOUNT = 0 then
		RSCHECK.close
		set RSCHECK = nothing
		response.redirect "setup.asp?id=3&err=2"
	else
		RSCHECK.close
		set RSCHECK = nothing
	end if

	RSDEPTSQL = "SELECT * FROM Departments ORDER BY SHORT"
	RSROOMSQL = "SELECT * FROM Rooms ORDER BY ROOMNO ASC"
	RSSTAFFSQL = "SELECT ID, FN, LN, CATEGORY, ENTITLEMENT, DEPT, DEFROOM, TITLE FROM Timetables"
				
	Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
	RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic

	Set RSROOM = Server.CreateObject("Adodb.RecordSet")
	RSROOM.Open RSROOMSQL, dataconn, adopenkeyset, adlockoptimistic

	Set RSSTAFF = Server.CreateObject("Adodb.RecordSet")
	RSSTAFF.Open RSSTAFFSQL, dataconn, adopenkeyset, adlockoptimistic
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li><a href="setup.asp?id=1">Welcome</a> :: </li>
			<li><a href="setup.asp?id=2">Step 1</a> :: </li>
			<li><a href="setup.asp?id=3">Step 2</a> :: </li>
			<li><a href="setup.asp?id=4"><b>Step 3</b></a> :: </li>
			<%
			RSUSERSQL = "SELECT * FROM Timetables"
			Set RSUSER = Server.CreateObject("Adodb.RecordSet")
			RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic

			if RSUSER.RECORDCOUNT => 1 then
			%>
			<li><a href="setup.asp?id=5">Step 4</a> :: </li>
			<%
			else
			%>
			<li>Step 4 :: </li>
			<%
			end if
			%>
			<li>Step 5 :: </li>
			<li>Step 6 :: </li>
			<li>Step 7 :: </li>
			<li>Complete</li>
			<%
			RSUSER.close
			set RSUSER = nothing
			%>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<form name="add" action="/pt/modules/ss/db/setup.asp?setuptype=3" method="post">
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Step 3</div>
		<div class="m_l_subtitle">Adding Staff</div>
		<div class="m_l_ins">This Step Is Important! Please Add All Required Staff With Care, And Then Click "Next".</div>
		<div class="m_l_sel" style="width: 100%;">
			<%
			if (request("err")) = "1" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">The Staff Details You Specified Already Exist!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "2" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Left Important Details Out! Please Try Again.</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "3" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Must Add At Least <b>One</b> Member Of Staff To Continue!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			else
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">Any Problems With Your Information Will Be Displayed Here.</td>
				</tr>
			</table>
			<hr size="1">
			<%
			end if
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('add');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="Add A Member Of Staff"></td>
					<td class="m_l_list_t"><b>Add A Member Of Staff</b></td>
				</tr>
				<tr id="list_add">
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
								<td colspan="2" class="m_l_list_t">
								<hr size="1">
								<div style="width: 100%; text-align: center;"><b><a href="#" onmouseup="document.add.submit();">Save</a></b> :: <b><a href="#" onmouseup="document.add.reset();">Clear</a></b></div>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep"></td>
				</tr>
			</table>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('added');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="Successfully Added Staff Members"></td>
					<td class="m_l_list_t"><b>Successfully Added Staff Members</b></td>
				</tr>
				<tr id="list_added" style="display: none;">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
						<%
						if RSSTAFF.RECORDCOUNT = 0 then
						%>
							<tr class="m_l_list_b" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);">
								<td class="m_l_list_t" style="height: 26px; padding-left: 5px;">No Staff Found!</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						<%
						else

						do until RSSTAFF.EOF
						
						RSDEPT2SQL = "SELECT * FROM Departments WHERE DEPTID = " & RSSTAFF("DEPT")
						
						Set RSDEPT2 = Server.CreateObject("Adodb.RecordSet")
						RSDEPT2.Open RSDEPT2SQL, dataconn, adopenkeyset, adlockoptimistic
						%>
							<tr class="m_l_list_b" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);" onmouseup="showdetail('<%=RSSTAFF("id")%>');">
								<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="<%=RSSTAFF("FN")%>&nbsp;<%=RSSTAFF("LN")%>"></td>
								<td class="m_l_list_t"><%=RSSTAFF("LN")%>, <%=Left(RSSTAFF("FN"),1)%>.</td>
							</tr>
							<tr id="list_<%=RSSTAFF("id")%>" style="display: none;">
								<td class="m_l_list_p"></td>
								<td class="m_l_list_t" style="line-height: 20px; padding-top: 7px;">
								<b>Full Name:</b> <%=RSSTAFF("Title")%>. <%=RSSTAFF("FN")%>&nbsp;<%=RSSTAFF("LN")%> (System ID No. <%=RSSTAFF("ID")%>)<br>
								<b>Department:</b> <%=RSDEPT2("FULL")%><br>
								<b>Position:</b>
								<%
								if RSSTAFF("CATEGORY") = "T" then
								response.write "Teacher"
								elseif RSSTAFF("CATEGORY") = "PT" then
								response.write "Principal Teacher"
								elseif RSSTAFF("CATEGORY") = "DHT" then
								response.write "Deputy Head Teacher"
								elseif RSSTAFF("CATEGORY") = "HT" then
								response.write "Head Teacher"
								elseif RSSTAFF("CATEGORY") = "PS" then
								response.write "Pupil Support"
								elseif RSSTAFF("CATEGORY") = "OC" then
								response.write "Outside Cover"
								else
								response.write "Unknown Position"
								end if
								%><br>
								<b>Usual Room:</b> 
								<%
								if RSSTAFF("DEFROOM") <> "" then
								response.write RSSTAFF("DEFROOM")
								else
								response.write "N/A"
								end if
								%><br>
								<b>Period Entitlement:</b> <%=RSSTAFF("ENTITLEMENT")%> Periods/Week
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						<%
						RSDEPT2.close
						set RSDEPT2 = nothing
				
						RSSTAFF.MOVENEXT
						loop

						end if
						%>
						</table>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep"></td>
				</tr>
			</table>
		<hr size="1">
		<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="setup.asp?id=5">Next></a></div>
		</div>
	</div>
	</form>
<%
	RSDEPT.close
	RSROOM.close
	RSSTAFF.close
	set RSDEPT = nothing
	set RSROOM = nothing
	set RSSTAFF = nothing

elseif pagetype = "5" then

	if session("whereiam") < 5 then
		response.redirect "setup.asp?err=1"
	else
		session("whereiam") = 5
	end if

	RSCHECKSQL = "SELECT * FROM Timetables"
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic

	if RSCHECK.RECORDCOUNT = 0 then
		RSCHECK.close
		set RSCHECK = nothing
		response.redirect "setup.asp?id=4&err=3"
	else
		RSCHECK.close
		set RSCHECK = nothing
	end if
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li><a href="setup.asp?id=1">Welcome</a> :: </li>
			<li><a href="setup.asp?id=2">Step 1</a> :: </li>
			<li><a href="setup.asp?id=3">Step 2</a> :: </li>
			<li><a href="setup.asp?id=4">Step 3</a> :: </li>
			<li><a href="setup.asp?id=5"><b>Step 4</b></a> :: </li>
			<li>Step 5 :: </li>
			<li>Step 6 :: </li>
			<li>Step 7 :: </li>
			<li>Complete</li>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Step 4a</div>
		<div class="m_l_subtitle">Weekends Or No Weekends?</div>
		<div class="m_l_ins">Does <%=var_est_short%> Use Weekends?</div>
		<form name="edit" action="/pt/modules/ss/db/setup.asp?setuptype=4" method="post">
		<div class="m_l_sel" style="width: 100%;">
		Here, Just Click "Yes" Or "No". This Can Be Changed Later By Logging In As An Administrator, And Going To<br>Settings > Change Period/Day Settings.
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr class="m_l_list_b">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_cog.gif" border="0" alt="Weekend Capability"></td>
					<td class="m_l_list_t"><b>Weekend Capability</b></td>
				</tr>
				<tr>
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr>
								<td class="m_l_list_t" style="padding-left: 5px;">Does <%=var_est_full%> Require The Use Of Weekends?</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;" onmouseup="document.edit.YN[0].checked = true; document.edit.submit();"><input type="radio" value="1" name="YN" id="YES"<%if var_est_enabled_weekends = "1" then%> checked<%else%><%end if%>><label for="YES">&nbsp;<b>Yes</b></label></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;" onmouseup="document.edit.YN[1].checked = true; document.edit.submit();"><label for="NO"><input type="radio" value="0" name="YN" id="NO"<%if var_est_enabled_weekends = "0" then%> checked<%else%><%end if%>>&nbsp;<b>No</b></label></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</div>
		</form>
	</div>
<%
elseif pagetype = "6" then

	if session("whereiam") < 5 then
		response.redirect "setup.asp?err=1"
	else
		session("whereiam") = 6
	end if
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li><a href="setup.asp?id=1">Welcome</a> :: </li>
			<li><a href="setup.asp?id=2">Step 1</a> :: </li>
			<li><a href="setup.asp?id=3">Step 2</a> :: </li>
			<li><a href="setup.asp?id=4">Step 3</a> :: </li>
			<li><a href="setup.asp?id=5"><b>Step 4</b></a> :: </li>
			<li>Step 5 :: </li>
			<li>Step 6 :: </li>
			<li>Step 7 :: </li>
			<li>Complete</li>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Step 4b</div>
		<div class="m_l_subtitle">Specifying How Many Periods Per Day</div>
		<div class="m_l_ins">How Many Periods A Day Are There?</div>
		<div class="m_l_sel" style="width: 100%;">
			<%
			if (request("ok")) = "1" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">
					<span style="color: #389F0A;"><b>Settings Successfully Applied</b></span>
					</td>
				</tr>
			</table>
			<hr size="1">
			<%
			else
			end if
			%>
			<b>The Highest Amount Of Periods This Version Goes To Is 10. If You Require More, Please Contact Server-ML, Who Will Be Able To Supply You
			With A Custom-Made Version.
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr class="m_l_list_b">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_cog.gif" border="0" alt="Periods In Each Day"></td>
					<td class="m_l_list_t"><b>Periods In Each Day</b></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
			</table>
			<%
			if var_est_enabled_weekends = "1" then
			%>
			<!--#include virtual="/pt/modules/ss/timetables/admin_setup_7day.inc"-->
			<%
			else
			%>
			<!--#include virtual="/pt/modules/ss/timetables/admin_setup_5day.inc"-->
			<%
			end if
			%>
		<hr size="1">
		<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="setup.asp?id=7">Next></a></div>
		</div>
	</div>
<%
elseif pagetype = "7" then

	if session("whereiam") < 6 then
		response.redirect "setup.asp?err=1"
	else
		session("whereiam") = 7
	end if

	if (request("part")) = "1" then
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li><a href="setup.asp?id=1">Welcome</a> :: </li>
			<li><a href="setup.asp?id=2">Step 1</a> :: </li>
			<li><a href="setup.asp?id=3">Step 2</a> :: </li>
			<li><a href="setup.asp?id=4">Step 3</a> :: </li>
			<li><a href="setup.asp?id=5">Step 4</a> :: </li>
			<li><a href="setup.asp?id=7"><b>Step 5</b></a> :: </li>
			<li>Step 6 :: </li>
			<li>Step 7 :: </li>
			<li>Complete</li>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Step 5</div>
		<div class="m_l_subtitle">Staff Timetables</div>
		<div class="m_l_ins">Please Read The Information Below, Then Click "Start".</div>
		<div class="m_l_sel" style="width: 100%;">
		<b>This Is A Time Consuming Step.</b><br>
		This Takes A While, Depending On How Many Staff There Are.<br>
		If You Click "Start", Everyone's Blank Timetable Will Be Displayed (Alphabetically), And You Will Have To Add In Their Details.<br>
		As This May Take A While, There Is An Option To Skip, Although Everyone Will Appear As "Free" And There Will Be No Classes To Cover, Until
		You Or Another Administrator Add In The Details!<br>
		<b>What Would You Like To Do?</b>
		<hr size="1">
		<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="setup.asp?id=8">Skip></a> :: <a href="setup.asp?id=7&amp;part=2">Start></a></div>
		</div>
	</div>
<%
	elseif (request("part")) = "2" then

	if(request("user")) = "" then
		currrecord = 0
	else
		currrecord = request("user")
	end if

	disprecord = currrecord + 1

	RSUSERSETUPSQL = "SELECT * FROM Timetables ORDER BY LN"
	
	Set RSUSERSETUP = Server.CreateObject("Adodb.RecordSet")
	RSUSERSETUP.Open RSUSERSETUPSQL, dataconn, adopenkeyset, adlockoptimistic

	RSUSERSETUP.MOVE(currrecord)
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li><a href="setup.asp?id=1">Welcome</a> :: </li>
			<li><a href="setup.asp?id=2">Step 1</a> :: </li>
			<li><a href="setup.asp?id=3">Step 2</a> :: </li>
			<li><a href="setup.asp?id=4">Step 3</a> :: </li>
			<li><a href="setup.asp?id=5">Step 4</a> :: </li>
			<li><a href="setup.asp?id=7"><b>Step 5</b></a> :: </li>
			<li>Step 6 :: </li>
			<li>Step 7 :: </li>
			<li>Complete</li>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Step 5</div>
		<div class="m_l_subtitle">Staff Timetables</div>
		<div class="m_l_ins">Displaying <%=RSUSERSETUP("FN")%>&nbsp;<%=RSUSERSETUP("LN")%>'s Timetable (No <%=disprecord%> Of <%=RSUSERSETUP.RECORDCOUNT%>)</div>
		<div class="m_l_sel" style="width: 100%;">
		To Edit A Period, Just Click Its Box, And A Popup Will Appear Allowing You To Add A Class To <%=RSUSERSETUP("FN")%>'s Timetable.<br>
		<%
		if disprecord = RSUSERSETUP.RECORDCOUNT then
		%>
		This Is The Last Timetable. When You Click "<b>Next</b>", You Will Be Taken To Step 6.
		<%
		else
		%>
		Once You Are Done With <%=RSUSERSETUP("FN")%>'s Timetable, Click "<b>Next Member</b>".
		<%
		end if
		%>
		<hr size="1">
			<%
			if var_est_enabled_weekends = "1" then
			%>
			<!--#include virtual="/pt/modules/ss/timetables/admin_setup_7day_stt.inc"-->
			<%
			else
			%>
			<!--#include virtual="/pt/modules/ss/timetables/admin_setup_5day_stt.inc"-->
			<%
			end if
			%>
		<hr size="1">
		<%
		if disprecord = RSUSERSETUP.RECORDCOUNT then
		%>
		<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="setup.asp?id=8">Next></a></div>
		<%
		else
		%>
		<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="setup.asp?id=7&amp;part=2&amp;user=<%=currrecord + 1%>">Next Member></a></div>
		<%
		end if
		%>
		</div>
	</div>
<%
	elseif (request("part")) = "3" then

		if (request("err")) = "1" then
%>
<div class="main" style="width: 634px;">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The <b>Initial Setup Account</b></div>
	<div class="smlb_r">
		<ul>
			<li><a href="#" onmouseup="self.close();" title="Click Here To Close This Popup Window.">Close Popup</a></li>
		</ul>
	</div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Staff Timetable</div>
		<div class="m_l_subtitle">You Forgot To Enter A Class Name!</div>
		<div class="m_l_ins">Please Go Back And Fill In The Field.</div>
		<div class="m_l_sel">
			Sorry, <%=session("sess_fn")%>, But You Have Forgotten To Enter The Class Name.<br>
			This Is Required In Order For The System To Determine A Member Of Staff Has A Class, So Therefore Is Very Important.<br>
			Please Go Back By Clicking "Return" Below And Fix The Problem.
			<div class="botopts">
				<ul>
					<li><a href="#" onmouseup="history.back();">Return</a></li>
					<li><a href="#" onmouseup="self.close();">Close The Popup Window</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	else
%>
<div class="main" style="width: 634px;">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The <b>Initial Setup Account</b></div>
	<div class="smlb_r">
		<ul>
			<li><a href="#" onmouseup="self.close();" title="Click Here To Close This Popup Window.">Close Popup</a></li>
		</ul>
	</div>
	<div class="topb_t"><%=var_pname%> Setup</div>


	<form name="edit" action="/pt/modules/ss/db/setup.asp?setuptype=6&amp;period=<%=request("period")%>&amp;day=<%=request("day")%>" method="post">
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Step 5</div>
		<div class="m_l_subtitle">Editing A Period</div>
		<div class="m_l_ins">Please Change The Settings Below For This Period, Then Click "Save".</div>
		<div class="m_l_sel">
		<%
		RSUSERSQL = "SELECT * FROM Timetables WHERE ID = " & request("user")
		RSROOMSQL = "SELECT * FROM Rooms"
		
		Set RSUSER = Server.CreateObject("Adodb.RecordSet")
		RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
		
		Set RSROOM = Server.CreateObject("Adodb.RecordSet")
		RSROOM.Open RSROOMSQL, dataconn, adopenkeyset, adlockoptimistic
		%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="<%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%>"></td>
					<td class="m_l_list_t"><b><%=RSUSER("LN")%>, <%=Left(RSUSER("FN"),1)%>.</b>, 
					<%
					if (request("day")) = "1" then
					response.write "Sunday"
					elseif (request("day")) = "2" then
					response.write "Monday"
					elseif (request("day")) = "3" then
					response.write "Tuesday"
					elseif (request("day")) = "4" then
					response.write "Wednesday"
					elseif (request("day")) = "5" then
					response.write "Thursday"
					elseif (request("day")) = "6" then
					response.write "Friday"
					elseif (request("day")) = "7" then
					response.write "Saturday"
					else
					end if
					%>, Period <%=(request("period"))%>
					</td>
				</tr>
				<tr>
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;" onmouseup="document.edit.ROOM.disabled=false; document.edit.NAME.disabled=false; document.edit.YN[0].checked = true;"><input type="radio" value="1" name="YN" id="YES"<%if RSUSER(request("period") & "_" & request("day")) <> "" then%> checked <%else%> <%end if%>onmouseup="document.edit.ROOM.disabled=false; document.edit.NAME.disabled=false;"><label for="YES">&nbsp;<b><%=RSUSER("FN")%> Has A Class This Period</b></label></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr>
								<td>
									<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
										<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
											<td style="width: 40px;"></td>
											<td class="m_l_list_t" style="height: 28px; width: 150px;">In Room</td>
											<td class="m_l_list_t" style="height: 28px;">
											<select class="b_std" size="1" name="ROOM"<%if RSUSER(request("period") & "_" & request("day")) <> "" then%><%else%> disabled<%end if%>>
											<%
											do until RSROOM.EOF
											%>
												<option<%if RSROOM("ROOMNO") = RSUSER("R" & request("period") & "_" & request("day") ) then%> selected<%else%><%end if%>><%=RSROOM("ROOMNO")%></option>
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
											<td style="width: 40px;"></td>
											<td class="m_l_list_t" style="height: 28px; width: 150px;">With A Class Name Of</td>
											<td class="m_l_list_t" style="height: 28px;"><input type="hidden" name="USER" value="<%=RSUSER("ID")%>"><input class="b_std" type="text" name="NAME" size="20" value="<%=RSUSER(request("period") & "_" & request("day"))%>"<%if RSUSER(request("period") & "_" & request("day")) <> "" then%><%else%> disabled<%end if%>></td>
										</tr>
									</table>							
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;" onmouseup="document.edit.ROOM.disabled='true'; document.edit.NAME.disabled='true'; document.edit.YN[1].checked = true;"><label for="NO" onmouseup="document.edit.ROOM.disabled='true'; document.edit.NAME.disabled='true';"><input type="radio" value="0" name="YN" id="NO"<%if RSUSER(request("period") & "_" & request("day")) <> "" then%><%else%> checked<%end if%>>&nbsp;<b><%=RSUSER("FN")%> Doesn't Has A Class This Period</b></label></td>
							</tr>
						</table>
						<hr size="1">
						<div style="width: 100%; text-align: center;"><b><a href="#" onmouseup="document.edit.submit();">Save</a></b> :: <b><a href="#" onmouseup="document.edit.reset();">Reset</a></b> :: <b><a href="#" onmouseup="self.close();">Close Popup</a></b></div>
					</td>
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
		response.redirect "setup.asp?id=7&part=1"
	end if

elseif pagetype = "8" then

	if session("whereiam") < 7 then
		response.redirect "setup.asp?err=1"
	else
		session("whereiam") = 8
	end if
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li><a href="setup.asp?id=1">Welcome</a> :: </li>
			<li><a href="setup.asp?id=2">Step 1</a> :: </li>
			<li><a href="setup.asp?id=3">Step 2</a> :: </li>
			<li><a href="setup.asp?id=4">Step 3</a> :: </li>
			<li><a href="setup.asp?id=5">Step 4</a> :: </li>
			<li><a href="setup.asp?id=7">Step 5</a> :: </li>
			<li><a href="setup.asp?id=8"><b>Step 6</b></a> :: </li>
			<li>Step 7 :: </li>
			<li>Complete</li>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Step 6</div>
		<div class="m_l_subtitle">Will Users Will Require A PIN?</div>
		<div class="m_l_ins">A PIN Is A Six-Digit Group Of Numbers Or Letters Users Can Enter As Well As A Password.</div>
		<form name="edit" action="/pt/modules/ss/db/setup.asp?setuptype=7" method="post">
		<div class="m_l_sel" style="width: 100%;">
		In Today's Security-Conscious World, People Want Their Data To Be Protected From Hackers And The Like.<br>
		This System Has A Built-In Feature Which Can Double The Security Of Passwords, Which Is A PIN.<br>
		It's A Six-Digit Group Of Numbers Or Letters Which Users Enter At The Login Screen As Well As Their Password.
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr class="m_l_list_b">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_cog.gif" border="0" alt="PINs"></td>
					<td class="m_l_list_t"><b>PINs</b></td>
				</tr>
				<tr>
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr>
								<td class="m_l_list_t" style="padding-left: 5px; line-height: 20px;">Does <%=var_est_full%> Need Extra Security By Means Of A PIN For Each User?<br>This Feature Can Be Turned On Later If It Is Not Required Now.</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;" onmouseup="document.edit.YN[0].checked = true; document.edit.submit();"><input type="radio" value="1" name="YN" id="YES"><label for="YES">&nbsp;<b>Yes</b></label></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;" onmouseup="document.edit.YN[1].checked = true; document.edit.submit();"><label for="NO"><input type="radio" value="0" name="YN" id="NO">&nbsp;<b>No</b></label></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</div>
		</form>
	</div>
<%
elseif pagetype = "9" then

	if session("whereiam") < 8 then
		response.redirect "setup.asp?err=1"
	else
		session("whereiam") = 9
	end if

	if (request("part")) = "1" then

	RSDEPTSQL = "SELECT * FROM Departments ORDER BY SHORT"
	RSROOMSQL = "SELECT * FROM Rooms ORDER BY ROOMNO ASC"
	RSSTAFFSQL = "SELECT ID, FN, LN, CATEGORY, ENTITLEMENT, DEPT, DEFROOM, TITLE FROM Timetables ORDER BY DEPT, LN"
				
	Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
	RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic

	Set RSROOM = Server.CreateObject("Adodb.RecordSet")
	RSROOM.Open RSROOMSQL, dataconn, adopenkeyset, adlockoptimistic

	Set RSSTAFF = Server.CreateObject("Adodb.RecordSet")
	RSSTAFF.Open RSSTAFFSQL, dataconn, adopenkeyset, adlockoptimistic
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li><a href="setup.asp?id=1">Welcome</a> :: </li>
			<li><a href="setup.asp?id=2">Step 1</a> :: </li>
			<li><a href="setup.asp?id=3">Step 2</a> :: </li>
			<li><a href="setup.asp?id=4">Step 3</a> :: </li>
			<li><a href="setup.asp?id=5">Step 4</a> :: </li>
			<li><a href="setup.asp?id=7">Step 5</a> :: </li>
			<li><a href="setup.asp?id=8">Step 6</a> :: </li>
			<li><a href="setup.asp?id=9"><b>Step 7</b></a> :: </li>
			<li>Complete</li>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Step 7</div>
		<div class="m_l_subtitle">Creating A Full-Access Admin Account</div>
		<div class="m_l_ins">Please Select The Member Of Staff That Will Get The Account From The List Below.</div>
		<div class="m_l_sel" style="width: 100%;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
						<%
						do until RSSTAFF.EOF
						
						RSDEPT2SQL = "SELECT * FROM Departments WHERE DEPTID = " & RSSTAFF("DEPT")
						
						Set RSDEPT2 = Server.CreateObject("Adodb.RecordSet")
						RSDEPT2.Open RSDEPT2SQL, dataconn, adopenkeyset, adlockoptimistic
						%>
							<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="location.href='setup.asp?id=9&amp;part=2&amp;user=<%=RSSTAFF("ID")%>';">
								<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="<%=RSSTAFF("FN")%>&nbsp;<%=RSSTAFF("LN")%>"></td>
								<td class="m_l_list_t">
								<b><%=RSSTAFF("LN")%>, <%=Left(RSSTAFF("FN"),1)%>.</b>
								(<%
								if RSSTAFF("CATEGORY") = "T" then
								response.write "Teacher"
								elseif RSSTAFF("CATEGORY") = "PT" then
								response.write "Principal Teacher"
								elseif RSSTAFF("CATEGORY") = "DHT" then
								response.write "Deputy Head Teacher"
								elseif RSSTAFF("CATEGORY") = "HT" then
								response.write "Head Teacher"
								elseif RSSTAFF("CATEGORY") = "PS" then
								response.write "Pupil Support"
								elseif RSSTAFF("CATEGORY") = "OC" then
								response.write "Outside Cover"
								else
								response.write "Unknown Position"
								end if
								%>, <%=RSDEPT2("FULL")%>)</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<%
							RSDEPT2.close
							set RSDEPT2 = nothing
				
							RSSTAFF.MOVENEXT
							loop
							%>
						</table>
		</div>
	</div>
<%
	RSDEPT.close
	RSROOM.close
	RSSTAFF.close
	set RSDEPT = nothing
	set RSROOM = nothing
	set RSSTAFF = nothing

	elseif (request("part")) = "2" then

		if (request("user")) <> "" then

		RSUSERSQL = "SELECT * FROM Timetables WHERE ID = " & request("user")

		Set RSUSER = Server.CreateObject("Adodb.RecordSet")
		RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic

		RSDEPTSQL = "SELECT * FROM Departments WHERE DEPTID = " & RSUSER("DEPT")

		Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
		RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li><a href="setup.asp?id=1">Welcome</a> :: </li>
			<li><a href="setup.asp?id=2">Step 1</a> :: </li>
			<li><a href="setup.asp?id=3">Step 2</a> :: </li>
			<li><a href="setup.asp?id=4">Step 3</a> :: </li>
			<li><a href="setup.asp?id=5">Step 4</a> :: </li>
			<li><a href="setup.asp?id=7">Step 5</a> :: </li>
			<li><a href="setup.asp?id=8">Step 6</a> :: </li>
			<li><a href="setup.asp?id=9"><b>Step 7</b></a> :: </li>
			<li>Complete</li>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Step 7</div>
		<div class="m_l_subtitle">Creating A Full-Access Admin Account</div>
		<div class="m_l_ins">Please Fill In Any Empty Fields Below, And Then Click "Create Account".</div>
		<form name="add" action="/pt/modules/ss/db/setup.asp?setuptype=9&amp;mode=1&amp;user=<%=request("user")%>" method="post">
		<div class="m_l_sel" style="width: 100%;">
			<%
			if (request("err")) = "1" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Left Out Some Information. All Information Is Required.</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "2" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">The Two Passwords You Entered Do Not Match. Please Try Again.</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "3" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">The User Account Does Not Exist.</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "4" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">The Username You Specified Already Exists. Please Try A Different <b>Username</b>.</td>
				</tr>
			</table>
			<hr size="1">
			<%
			else
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">Any Problems With Your Information Will Be Displayed Here.</td>
				</tr>
			</table>
			<hr size="1">
			<%
			end if
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_t" style="height: 28px; width: 120px; padding-left: 5px;"><b>Name</b></td>
					<td class="m_l_list_t"><%=RSUSER("TITLE")%>. <%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_t" style="height: 28px; width: 120px; padding-left: 5px;"><b>Department</b></td>
					<td class="m_l_list_t"><%=RSDEPT("FULL")%></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_t" style="height: 28px; width: 120px; padding-left: 5px;"><b>Username</b></td>
					<td class="m_l_list_t"><input class="b_std" type="text" name="UN" size="20" value="<%=RSUSER("LN")%><%=left(RSUSER("FN"),1)%>">&nbsp;The Initial Value Is A Suggestion - You Can Change It If You Want!</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_t" style="height: 28px; width: 120px; padding-left: 5px;"><b>Password</b></td>
					<td class="m_l_list_t"><input class="b_pwd" type="password" name="PW" size="20"></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_t" style="height: 28px; width: 120px; padding-left: 5px;"><b>Verify Password</b></td>
					<td class="m_l_list_t"><input class="b_pwd" type="password" name="PWV" size="20"></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
				if var_est_enabled_pin = 1 then
				%>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_t" style="height: 28px; width: 120px; padding-left: 5px;"><b>PIN</b></td>
					<td class="m_l_list_t" style="font-size: 10pt;"><input class="b_pwd" type="password" name="p1" size="3" maxlength="3" onkeyup="pjump();">&nbsp;/&nbsp;<input class="b_pwd" type="password" name="p2" size="3" maxlength="3"></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
				else
				end if
				%>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_t" style="height: 28px; width: 120px; padding-left: 5px;"><b>E-Mail Address</b></td>
					<td class="m_l_list_t"><input class="b_std" type="text" name="EM" size="20">&nbsp;@<%=var_emaildomain1%></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
			</table>
			<%
			if var_est_enabled_pin <> 1 then
			%>
			<div style="padding-left: 5px;">
			PINs Are Currently Disabled. There Is No Need To Enter One Right Now, But If They Are Enabled At A Later Stage, Your PIN Will Be Blank.
			You Can Then Create A PIN From "Change My Details" At The Top Right Of Each Page.
			</div>
			<%
			else
			end if
			%>
		<hr size="1">
		<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="#" onmouseup="document.add.submit();">Create Account></a></div>
		</div>
	</div>
	</form>
<%
		RSDEPT.close
		set RSDEPT = nothing

		RSUSER.close
		set RSUSER = nothing

		else
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li><a href="setup.asp?id=1">Welcome</a> :: </li>
			<li><a href="setup.asp?id=2">Step 1</a> :: </li>
			<li><a href="setup.asp?id=3">Step 2</a> :: </li>
			<li><a href="setup.asp?id=4">Step 3</a> :: </li>
			<li><a href="setup.asp?id=5">Step 4</a> :: </li>
			<li><a href="setup.asp?id=7">Step 5</a> :: </li>
			<li><a href="setup.asp?id=8">Step 6</a> :: </li>
			<li><a href="setup.asp?id=9"><b>Step 7</b></a> :: </li>
			<li>Complete</li>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<form name="add" action="/pt/modules/ss/db/setup.asp?setuptype=9&amp;mode=2" method="post">
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Step 7</div>
		<div class="m_l_subtitle">Creating A Full-Access Admin Account</div>
		<div class="m_l_ins">Please Fill In Any Empty Fields Below, And Then Click "Create Account".</div>
		<div class="m_l_sel" style="width: 100%;">
			<%
			if (request("err")) = "1" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Left Out Some Information. All Information Is Required.</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "2" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">The Two Passwords You Entered Do Not Match. Please Try Again.</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "3" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">The Username You Specified Already Exists. Please Try A Different <b>Username</b>.</td>
				</tr>
			</table>
			<hr size="1">
			<%
			else
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">Any Problems With Your Information Will Be Displayed Here.</td>
				</tr>
			</table>
			<hr size="1">
			<%
			end if
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_t" style="height: 28px; width: 120px; padding-left: 5px;"><b>Title</b></td>
					<td class="m_l_list_t">
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
					<td class="m_l_list_t" style="height: 28px; width: 120px; padding-left: 5px;"><b>First Name</b></td>
					<td class="m_l_list_t"><input class="b_std" type="text" name="FN" size="20"></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_t" style="height: 28px; width: 120px; padding-left: 5px;"><b>Last Name</b></td>
					<td class="m_l_list_t"><input class="b_std" type="text" name="LN" size="20"></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_t" style="height: 28px; width: 120px; padding-left: 5px;"><b>Username</b></td>
					<td class="m_l_list_t"><input class="b_std" type="text" name="UN" size="20"></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_t" style="height: 28px; width: 120px; padding-left: 5px;"><b>Password</b></td>
					<td class="m_l_list_t"><input class="b_pwd" type="password" name="PW" size="20"></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_t" style="height: 28px; width: 120px; padding-left: 5px;"><b>Verify Password</b></td>
					<td class="m_l_list_t"><input class="b_pwd" type="password" name="PWV" size="20"></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
				if var_est_enabled_pin = 1 then
				%>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_t" style="height: 28px; width: 120px; padding-left: 5px;"><b>PIN</b></td>
					<td class="m_l_list_t" style="font-size: 10pt;"><input class="b_pwd" type="password" name="p1" size="3" maxlength="3" onkeyup="pjump();">&nbsp;/&nbsp;<input class="b_pwd" type="password" name="p2" size="3" maxlength="3"></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
				else
				end if
				%>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_t" style="height: 28px; width: 120px; padding-left: 5px;"><b>E-Mail Address</b></td>
					<td class="m_l_list_t"><input class="b_std" type="text" name="EM" size="20">&nbsp;@<%=var_emaildomain1%></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
			</table>
			<%
			if var_est_enabled_pin <> 1 then
			%>
			<div style="padding-left: 5px;">
			PINs Are Currently Disabled. There Is No Need To Enter One Right Now, But If They Are Enabled At A Later Stage, Your PIN Will Be Blank.
			You Can Then Create A PIN From "Change My Details" At The Top Right Of Each Page.
			</div>
			<%
			else
			end if
			%>
		<hr size="1">
		<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="#" onmouseup="document.add.submit();">Create Account></a></div>
		</div>
	</div>
	</form>
<%
		end if

	else
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li><a href="setup.asp?id=1">Welcome</a> :: </li>
			<li><a href="setup.asp?id=2">Step 1</a> :: </li>
			<li><a href="setup.asp?id=3">Step 2</a> :: </li>
			<li><a href="setup.asp?id=4">Step 3</a> :: </li>
			<li><a href="setup.asp?id=5">Step 4</a> :: </li>
			<li><a href="setup.asp?id=7">Step 5</a> :: </li>
			<li><a href="setup.asp?id=8">Step 6</a> :: </li>
			<li><a href="setup.asp?id=9"><b>Step 7</b></a> :: </li>
			<li>Complete</li>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<form name="edit" action="/pt/modules/ss/db/setup.asp?setuptype=8" method="post">
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Step 7</div>
		<div class="m_l_subtitle">Creating A Full-Access Admin Account</div>
		<div class="m_l_ins">Create A Full-Access Admin Account.</div>
		<div class="m_l_sel" style="width: 100%;">
		This System Has Two Levels Of Administration Accounts: Full-Access And Partial Access.<br>
		Full Access Is Given To Staff Who Are Trusted With Managing All Aspects Of The System.<br>
		Partial Access Accounts Are Given To Heads Of Each Department So They May Sort Out Cover For Their Absent Staff.<br><br>
		A Full-Access Account Needs To Be Created So Changes And Maintenance Can Be Performed.
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr class="m_l_list_b">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="Full-Access Administrator"></td>
					<td class="m_l_list_t"><b>Full-Access Administrator</b></td>
				</tr>
				<tr>
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr>
								<td class="m_l_list_t" style="padding-left: 5px;">Although More Full-Access Accounts Can Be Created Later, Will This Initial Account Belong To A Member Of Teaching Staff?</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;" onmouseup="document.edit.YN[0].checked = true; document.edit.submit();"><input type="radio" value="1" name="YN" id="YES"><label for="YES">&nbsp;<b>Yes</b></label></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;" onmouseup="document.edit.YN[1].checked = true; document.edit.submit();"><label for="NO"><input type="radio" value="0" name="YN" id="NO">&nbsp;<b>No</b></label></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</div>
	</div>
	</form>
<%
	end if

elseif pagetype = "10" then

	if session("whereiam") < 10 then
		response.redirect "setup.asp?err=1"
	else
		session("whereiam") = 10
	end if
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li>Welcome :: </li>
			<li>Step 1 :: </li>
			<li>Step 2 :: </li>
			<li>Step 3 :: </li>
			<li>Step 4 :: </li>
			<li>Step 5 :: </li>
			<li>Step 6 :: </li>
			<li><b>Step 7</b> :: </li>
			<li>Complete</li>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Please Wait...</div>
		<div class="m_l_subtitle">Settings Are Being Applied...</div>
		<div class="m_l_ins">Please Wait A Moment While All The Settings Are Saved...</div>
		<div class="m_l_sel" style="width: 100%;">
			<b>Please DO NOT Close This Window, You May Loose All The Settings You Defined!</b><br>
			<div style="text-align: center;"><img src="/pt/media/admin/progress.gif" border="0" alt="Progress Boxes, A Fresh Alternative From A Bar."></div>
		</div>
	</div>
<%
elseif pagetype = "11" then

	if session("whereiam") < 11 then
		response.redirect "setup.asp?err=1"
	else
		session("whereiam") = 11
	end if
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li>Welcome :: </li>
			<li>Step 1 :: </li>
			<li>Step 2 :: </li>
			<li>Step 3 :: </li>
			<li>Step 4 :: </li>
			<li>Step 5 :: </li>
			<li>Step 6 :: </li>
			<li>Step 7 :: </li>
			<li><b>Complete</b></li>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Thank You!</div>
		<div class="m_l_subtitle">The System Is Ready To Go!</div>
		<div class="m_l_ins">All Information You Supplied Has Been Successfully Applied!</div>
		<div class="m_l_sel" style="width: 100%;">
		Thank You Very Much Indeed For Setting Up The System.<br>
		It Is Now Ready To Run And Accept User Signups.<br>
		You Can Also Login To The Admin Using The Full-Access Account That You Just Created.<br><br>
		<b>Thank You And We Hope That The PleaseTakes System Runs As You Expect It To, And Thank You For Choosing Server-ML!</b><br>
		To Continue, Click "Logout" Below And You Will Be Able To Login With The Full-Access Admin Account.
		<hr size="1">
		<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="#" onmouseup="location.href='/pt/modules/ss/usersys/logout.asp?id=2'">Logout></a></div>
		</div>
	</div>
<%
elseif (request("err")) = "1" then
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li>Welcome :: </li>
			<li>Step 1 :: </li>
			<li>Step 2 :: </li>
			<li>Step 3 :: </li>
			<li>Step 4 :: </li>
			<li>Step 5 :: </li>
			<li>Step 6 :: </li>
			<li>Step 7 :: </li>
			<li>Complete</li>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Sorry!</div>
		<div class="m_l_subtitle">Please Don't Jump The Queue!</div>
		<div class="m_l_ins">Please Don't Skip...It'll Ruin All The Information You Enter!</div>
		<div class="m_l_sel" style="width: 100%;">
		This Wizard Was Ordered For A Reason, So Succeeding Steps Have The Necessary Information To Work.<br>
		<b>Please Don't Change Querystring Details!</b><br><br>
		Please Click The Self-Explanatory Link Below To Continue.
		<hr size="1">
		<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="setup.asp?id=<%=session("whereiam")%>">Go To Where I SHOULD Be></a></div>
		</div>
	</div>
<%
elseif (request("err")) = "2" then
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li>Welcome :: </li>
			<li>Step 1 :: </li>
			<li>Step 2 :: </li>
			<li>Step 3 :: </li>
			<li>Step 4 :: </li>
			<li>Step 5 :: </li>
			<li>Step 6 :: </li>
			<li>Step 7 :: </li>
			<li>Complete</li>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Sorry!</div>
		<div class="m_l_subtitle">Access Is Denied!</div>
		<div class="m_l_ins">Access To The Page You Requested Has Been Denied.</div>
		<div class="m_l_sel" style="width: 100%;">
		Sorry, But This Page Is Temporarily Unavailable Due To The Fact That The Setup Wizard Has Not Been Successfully Completed.<br>
		Please Complete The Wizard And Login Using A Full-Access Admin Account, And Try Again.
		<hr size="1">
		<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="setup.asp?id=<%=session("whereiam")%>">Go To The Setup Wizard></a></div>
		</div>
	</div>
<%
else
%>
<div class="main">
	<div class="smlb_k"></div>
	<div class="smlb_l">This Is The Initial <b>Setup Account</b></div>
	<div class="topb_t"><%=var_pname%> Setup</div>
	<div class="topb_m">
		<ul>
			<li><a href="setup.asp?id=1"><b>Welcome</b></a> :: </li>
			<li><a href="setup.asp?id=2">Step 1</a> :: </li>
			<li>Step 2 :: </li>
			<li>Step 3 :: </li>
			<li>Step 4 :: </li>
			<li>Step 5 :: </li>
			<li>Step 6 :: </li>
			<li>Step 7 :: </li>
			<li>Complete</li>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title"><script language="javascript" type="text/javascript">document.write(daymsg)</script>!</div>
		<div class="m_l_subtitle">Welcome To The System!</div>
		<div class="m_l_ins">Thank You For Choosing Server-ML!</div>
		<div class="m_l_sel" style="width: 100%;">
		Hello, And Welcome To The Server-ML.co.uk PleaseTakes System!<br>
		We Would Like To Thank You For Choosing Server-ML Products, And Hope That You Have An Error-Free And Most Of All Enjoyable Time Using This System.<br>
		However, Before You And Your Establishment May Use This System, It Must First Be Setup With The Required Details.<br>
		This Is What This Wizard Helps You Do: Set Up The System.<br>
		<b>To Begin, Just Click "Start" Below.</b>
		<hr size="1">
		<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="setup.asp?id=2">Start></a></div>
		</div>
	</div>
<%
end if
%>

</div>

</body>

</html>

<%
end if
%><!--#include virtual="/pt/modules/ss/p_e.inc"-->