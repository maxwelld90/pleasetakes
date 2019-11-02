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

<body>

<div class="smlb_b"></div>
<div class="topb_b"></div>

<div class="main">
	<!--#include virtual="/pt/modules/ss/topbar/admin.inc"-->
<%
if pagetype = "2" then

	if (request("room")) <> "" then

	RSROOMSQL = "SELECT * FROM Rooms Where ID = " & (request("room"))
	Set RSROOM = Server.CreateObject("Adodb.RecordSet")
	RSROOM.Open RSROOMSQL, dataconn, adopenkeyset, adlockoptimistic

		if RSROOM.RECORDCOUNT = 0 then
%>
	<div class="m_l">
		<div class="m_l_title">Sorry!</div>
		<div class="m_l_subtitle">The Room You Specified Cannot Be Found!</div>
		<div class="m_l_ins">Please Go Back And Try Again.</div>
		<div class="m_l_sel">
			Sorry <%=session("sess_fn")%>, But The Room That You Specifed Cannot Be Found In The Database.<br>
			Please Go Back And Try Again.
			
			<div class="botopts">
				<ul>
					<li><a href="#" onmouseup="history.back();">Go Back</a></li>
					<li><a href="rooms.asp?id=1">Return To The Room Management Homepage</a></li>
					<li><a href="default.asp?id=1">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	else
%>
	<form name="edit" action="/pt/modules/ss/db/edit.asp?edittype=8&amp;roomid=<%=RSROOM("ID")%>" method="post">
	<div class="m_l">
		<div class="m_l_title">Edit A Room</div>
		<div class="m_l_subtitle">Edit Room 's Name/Number</div>
		<div class="m_l_ins">Please Enter The New Name And/Or Number Below, And Click "Save".</div>
		<div class="m_l_sel">
			<%
			if (request("err")) = "1" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Must Specify A New Name/Number For The Room!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "2" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">The Room You Specified Already Exists, So Editing Failed!</td>
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
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('edit');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_room.gif" border="0" alt="Question Mark"></td>
					<td class="m_l_list_t"><b>Edit A Room</b></td>
				</tr>
				<tr id="list_edit">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">

						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 140px; padding-left: 5px;"><b>Room Name/Number</b></td>
								<td ><input class="b_std" type="text" name="ROOM" size="20" value="<%=RSROOM("ROOMNO")%>"></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2" class="m_l_list_t">
								<hr size="1">
								<div style="width: 100%; text-align: center;"><b><a href="#" onmouseup="document.edit.submit();">Save</a></b> :: <b><a href="#" onmouseup="document.edit.reset();">Reset</a></b></div>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep"></td>
				</tr>
			</table>
		</div>
	</div>
	</form>
<%

		RSROOM.close
		set RSROOM = nothing

		end if

	else

	if (request("gd")) = "1" then
%>
	<div class="m_l">
		<div class="m_l_title">Add A Room</div>
		<div class="m_l_subtitle">Congratulations, Room Added!</div>
		<div class="m_l_ins">You Have Successfully Added A Room To The System!</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Congratulations, <%=session("sess_fn")%>! You Have Successfully Added A Room.<br>
			What Do You Want To Do Now?
			<div class="botopts">
				<ul>
					<li><a href="rooms.asp?id=2">Add Another Room</a></li>
					<li><a href="rooms.asp?id=1">Return To Room Management</a></li>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	elseif (request("gd")) = "2" then
%>
	<div class="m_l">
		<div class="m_l_title">Edit A Room</div>
		<div class="m_l_subtitle">Congratulations, Room Successfully Edited!</div>
		<div class="m_l_ins">You Have Successfully Edited The Room's Details!</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Congratulations, <%=session("sess_fn")%>! You Have Successfully Edited The Rooms Name And/Or Number.
			This Change Comes Into Immediate Effect.<br>
			What Do You Want To Do Now?
			<div class="botopts">
				<ul>
					<li><a href="rooms.asp?id=3">Edit Another Room</a></li>
					<li><a href="rooms.asp?id=2">Add A Room</a></li>
					<li><a href="rooms.asp?id=1">Return To Room Management</a></li>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	else
%>
	<form name="add" action="/pt/modules/ss/db/add.asp?addtype=4" method="post">
	<div class="m_l">
		<div class="m_l_title">Add A Room</div>
		<div class="m_l_subtitle">Add A Room To The System.</div>
		<div class="m_l_ins">Please Enter The Room Number In Below, And Click "Save".</div>
		<div class="m_l_sel">
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
					<td class="m_l_list_p"><img src="/pt/media/icons/16_room.gif" border="0" alt="Question Mark"></td>
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
		</div>
	</div>
	</form>
<%
	end if

	end if

elseif pagetype = "3" then
%>
	<div class="m_l">
		<div class="m_l_title">Edit A Room's Details</div>
		<div class="m_l_subtitle">Edit A Room's Name And/Or Number.</div>
		<div class="m_l_ins">Please Click The Room You Wish To Edit.</div>
		<div class="m_l_sel">
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('added');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_room.gif" border="0" alt="Question Mark"></td>
					<td class="m_l_list_t"><b>Which Room Do You Wish To Edit?</b></td>
				</tr>
				<tr id="list_added">
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
								<td class="m_l_list_t" style="height: 26px; padding-left: 5px;">No Rooms Found! Want To <a href="rooms.asp?id=2">Add</a> One?</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						<%
						else

						do until RSROOM.EOF
						%>
							<tr class="m_l_list_b" >
								<td class="m_l_list_t" style="width: 200px; height: 26px; padding-left: 5px;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);" onmouseup="location.href='rooms.asp?id=2&amp;room=<%=RSROOM("ID")%>'"><%=RSROOM("ROOMNO")%></td>
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
								<td class="m_l_list_t" style="width: 232px; height: 26px; padding-left: 5px;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);" onmouseup="location.href='rooms.asp?id=2&amp;room=<%=RSROOM("ID")%>'"><%=RSROOM("ROOMNO")%></td>
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
								<td class="m_l_list_t" style="width: 232px; height: 26px; padding-left: 5px;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);" onmouseup="location.href='rooms.asp?id=2&amp;room=<%=RSROOM("ID")%>'"><%=RSROOM("ROOMNO")%></td>
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
						set RSROOM = nothing
						%>
							</table>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep"></td>
				</tr>
			</table>
		</div>
	</div>
<%
elseif pagetype = "4" then

	if (request("gd")) = "1" then
%>
	<div class="m_l">
		<div class="m_l_title">Congratulations!</div>
		<div class="m_l_subtitle">The Room You Selected Has Been Removed!</div>
		<div class="m_l_ins">Now Please Choose An Option From Below.</div>
		<div class="m_l_sel">
			Congratulations, <%=session("sess_fn")%>! The Room You Asked To Delete Has Been Removed Successfully.<br>
			Please Choose An Option From Below To Continue.			
			<div class="botopts">
				<ul>
					<li><a href="rooms.asp?id=4">Delete Another Room</a></li>
					<li><a href="rooms.asp?id=2">Add A Room</a></li>
					<li><a href="rooms.asp?id=1">Return To The Room Management Homepage</a></li>
					<li><a href="default.asp?id=1">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	elseif (request("err")) = "1" then
%>
	<div class="m_l">
		<div class="m_l_title">Sorry!</div>
		<div class="m_l_subtitle">The Room You Specified Cannot Be Found!</div>
		<div class="m_l_ins">Now Please Choose An Option From Below.</div>
		<div class="m_l_sel">
			Sorry, <%=session("sess_fn")%>, But The System Is Unable To Find The Room You Specified.<br>
			Please Go Back And Try Again.			
			<div class="botopts">
				<ul>
					<li><a href="#" onmouseup="history.back();">Go Back</a></li>
					<li><a href="rooms.asp?id=4">Delete Another Room</a></li>
					<li><a href="rooms.asp?id=2">Add A Room</a></li>
					<li><a href="rooms.asp?id=1">Return To The Room Management Homepage</a></li>
					<li><a href="default.asp?id=1">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	else
%>
	<div class="m_l">
		<div class="m_l_title">Delete A Room</div>
		<div class="m_l_subtitle">Remove A Room From The System's Database.</div>
		<div class="m_l_ins">Please Select The Room You Wish To Delete. There Is No Conformation.</div>
		<div class="m_l_sel">
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('added');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_room.gif" border="0" alt="Keys To A Room"></td>
					<td class="m_l_list_t"><b>Which Room Do You Wish To Delete?</b></td>
				</tr>
				<tr id="list_added">
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
								<td class="m_l_list_t" style="height: 26px; padding-left: 5px;">No Rooms Found! Want To <a href="rooms.asp?id=2">Add</a> One?</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						<%
						else

						do until RSROOM.EOF
						%>
							<tr class="m_l_list_b" >
								<td class="m_l_list_t" style="width: 200px; height: 26px; padding-left: 5px;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);" onmouseup="location.href='/pt/modules/ss/db/delete.asp?deltype=3&amp;room=<%=RSROOM("ID")%>'"><%=RSROOM("ROOMNO")%></td>
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
								<td class="m_l_list_t" style="width: 232px; height: 26px; padding-left: 5px;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);" onmouseup="location.href='/pt/modules/ss/db/delete.asp?deltype=3&amp;room=<%=RSROOM("ID")%>'"><%=RSROOM("ROOMNO")%></td>
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
								<td class="m_l_list_t" style="width: 232px; height: 26px; padding-left: 5px;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);" onmouseup="location.href='/pt/modules/ss/db/delete.asp?deltype=3&amp;room=<%=RSROOM("ID")%>'"><%=RSROOM("ROOMNO")%></td>
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
						set RSROOM = nothing
						%>
							</table>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep"></td>
				</tr>
			</table>
		</div>
	</div>
<%
	end if

elseif pagetype = "5" then

	if (request("room")) <> "" then

		RSROOMSQL = "SELECT * FROM Rooms WHERE ID = " & request("room")
		Set RSROOM = Server.CreateObject("Adodb.RecordSet")
		RSROOM.Open RSROOMSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<div class="m_l">
		<div class="m_l_title">Actions Available</div>
		<div class="m_l_subtitle">What You Can Do To This Room</div>
		<div class="m_l_ins">Available Actions For Room <%=RSROOM("ROOMNO")%> Are Listed Below.</div>
		<div class="m_l_sel">
			<div class="botopts">
				<ul>
					<li><a href="rooms.asp?id=2&amp;room=<%=RSROOM("ID")%>">Edit The Room's Name/Number</a></li>
					<li><a href="/pt/modules/ss/db/delete.asp?deltype=3&amp;room=<%=RSROOM("ID")%>">Delete The Room (No Conformation)</a><p></li>
					<li><a href="rooms.asp?id=5">Return To Viewing All The Rooms</a></li>
					<li><a href="rooms.asp?id=1">Return To The Room Management Homepage</a></li>
					<li><a href="default.asp?id=1">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
		RSROOM.close
		set RSROOM = nothing

	else
%>
	<div class="m_l">
		<div class="m_l_title">View Current Rooms</div>
		<div class="m_l_subtitle">Shows You All The Rooms Being Used In The System.</div>
		<div class="m_l_ins">To View Any Actions You Can Do With A Room, Just Click It.</div>
		<div class="m_l_sel">
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('added');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_room.gif" border="0" alt="Keys To A Room"></td>
					<td class="m_l_list_t"><b>Displaying All Rooms</b></td>
				</tr>
				<tr id="list_added">
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
								<td class="m_l_list_t" style="height: 26px; padding-left: 5px;">No Rooms Found! Want To <a href="rooms.asp?id=2">Add</a> One?</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						<%
						else

						do until RSROOM.EOF
						%>
							<tr class="m_l_list_b" >
								<td class="m_l_list_t" style="width: 200px; height: 26px; padding-left: 5px;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);" onmouseup="location.href='rooms.asp?id=5&amp;room=<%=RSROOM("ID")%>'"><%=RSROOM("ROOMNO")%></td>
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
								<td class="m_l_list_t" style="width: 232px; height: 26px; padding-left: 5px;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);" onmouseup="location.href='rooms.asp?id=5&amp;room=<%=RSROOM("ID")%>'"><%=RSROOM("ROOMNO")%></td>
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
								<td class="m_l_list_t" style="width: 232px; height: 26px; padding-left: 5px;" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);" onmouseup="location.href='rooms.asp?id=5&amp;room=<%=RSROOM("ID")%>'"><%=RSROOM("ROOMNO")%></td>
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
						set RSROOM = nothing
						%>
							</table>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep"></td>
				</tr>
			</table>
		</div>
	</div>
<%
	end if
else
%>
	<div class="m_l">
		<div class="m_l_title">Room Management</div>
		<div class="m_l_subtitle">Add Or Remove Rooms From The System.</div>
		<div class="m_l_ins">Please Choose A Task From Below...</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_t_l"><a href="rooms.asp?id=2"><img src="/pt/media/icons/48_room.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="rooms.asp?id=2">Add A Room</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="rooms.asp?id=3"><img src="/pt/media/icons/48_edit.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="rooms.asp?id=3">Edit A Room's Details</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="rooms.asp?id=4"><img src="/pt/media/icons/48_del.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="rooms.asp?id=4">Delete A Room</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="rooms.asp?id=5"><img src="/pt/media/icons/48_room.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="rooms.asp?id=5">View Current Rooms</a></td>
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