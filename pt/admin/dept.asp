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

	if (request("dept")) <> "" then

	RSDEPTSQL = "SELECT * FROM Departments Where ID = " & (request("dept"))
	Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
	RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic

		if RSDEPT.RECORDCOUNT = 0 then
%>
	<div class="m_l">
		<div class="m_l_title">Sorry!</div>
		<div class="m_l_subtitle">The Department You Specified Cannot Be Found!</div>
		<div class="m_l_ins">Please Go Back And Try Again.</div>
		<div class="m_l_sel">
			Sorry <%=session("sess_fn")%>, But The Department That You Specifed Cannot Be Found In The Database.<br>
			Please Go Back And Try Again.
			
			<div class="botopts">
				<ul>
					<li><a href="#" onmouseup="history.back();">Go Back</a></li>
					<li><a href="dept.asp?id=1">Return To The Department Management Homepage</a></li>
					<li><a href="default.asp?id=1">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	else
%>
	<form name="edit" action="/pt/modules/ss/db/edit.asp?edittype=10&amp;deptid=<%=RSDEPT("ID")%>" method="post">
	<div class="m_l">
		<div class="m_l_title">Edit A Department</div>
		<div class="m_l_subtitle">Edit A Department's Details</div>
		<div class="m_l_ins">Please Enter The New Details For The Department, And Click "Save".</div>
		<div class="m_l_sel">
			<%
			if (request("err")) = "1" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Must Specify New Details For The Department!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "2" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">The Department You Specified Already Exists, So Editing Failed!</td>
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
					<td class="m_l_list_p"><img src="/pt/media/icons/16_dept.gif" border="0" alt="Edit A Department"></td>
					<td class="m_l_list_t"><b>Edit A Department</b></td>
				</tr>
				<tr id="list_edit">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">

						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 140px; padding-left: 5px;"><b>Full Name</b></td>
								<td ><input class="b_std" type="text" name="FULL" size="20" value="<%=RSDEPT("FULL")%>"></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 140px; padding-left: 5px;"><b>Short Name</b></td>
								<td ><input class="b_std" type="text" name="SHORT" size="20" value="<%=RSDEPT("SHORT")%>"></td>
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

		RSDEPT.close
		set RSDEPT = nothing

		end if

	else

	if (request("gd")) = "1" then
%>
	<div class="m_l">
		<div class="m_l_title">Add A Room</div>
		<div class="m_l_subtitle">Congratulations, Department Added!</div>
		<div class="m_l_ins">You Have Successfully Added A Department To The System!</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Congratulations, <%=session("sess_fn")%>! You Have Successfully Added A Department.<br>
			It Can Now Be Used With Staff Members.<br>
			What Do You Want To Do Now?
			<div class="botopts">
				<ul>
					<li><a href="dept.asp?id=2">Add Another Department</a></li>
					<li><a href="dept.asp?id=1">Return To Department Management</a></li>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	elseif (request("gd")) = "2" then
%>
	<div class="m_l">
		<div class="m_l_title">Edit A Department</div>
		<div class="m_l_subtitle">Congratulations, Department Successfully Edited!</div>
		<div class="m_l_ins">You Have Successfully Edited The Department's Details!</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Congratulations, <%=session("sess_fn")%>! You Have Successfully Edited The Department's Details.
			This Change Comes Into Immediate Effect.<br>
			What Do You Want To Do Now?
			<div class="botopts">
				<ul>
					<li><a href="dept.asp?id=3">Edit Another Department</a></li>
					<li><a href="dept.asp?id=2">Add A Department</a></li>
					<li><a href="dept.asp?id=1">Return To Department Management</a></li>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	else
%>
	<form name="add" action="/pt/modules/ss/db/add.asp?addtype=7" method="post">
	<div class="m_l">
		<div class="m_l_title">Add A Department</div>
		<div class="m_l_subtitle">Add A Department To The System.</div>
		<div class="m_l_ins">Please Enter The New Department's Details In Below, And Click "Save".</div>
		<div class="m_l_sel">
			<%
			if (request("err")) = "1" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Must Specify At Least A Full Name For The Department You Wish To Add!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "2" then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">The Department You Specified Already Exists!</td>
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
					<td class="m_l_list_p"><img src="/pt/media/icons/16_dept.gif" border="0" alt="Add A Department"></td>
					<td class="m_l_list_t"><b>Add A Department</b></td>
				</tr>
				<tr id="list_add">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">

						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 140px; padding-left: 5px;"><b>Full Name</b></td>
								<td ><input class="b_std" type="text" name="FULL" size="20"></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 140px; padding-left: 5px;"><b>Short Name</b></td>
								<td><input class="b_std" type="text" name="SHORT" size="20"></td>
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
						<br>
						The "Short Name" Field Is Optional, The "Full Name" Will Be Used Instead If No Short Name Is Supplied.
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
		<div class="m_l_title">Edit A Department's Details</div>
		<div class="m_l_subtitle">Edit A Department's Details</div>
		<div class="m_l_ins">Please Click The Department You Wish To Edit.</div>
		<div class="m_l_sel">
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('added');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_dept.gif" border="0" alt="Which Department Do You Wish To Edit?"></td>
					<td class="m_l_list_t"><b>Which Department Do You Wish To Edit?</b></td>
				</tr>
				<tr id="list_added">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">

						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
						<%
						RSDEPTSQL = "SELECT * FROM Departments ORDER BY SHORT ASC"
						Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
						RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic

						if RSDEPT.RECORDCOUNT = 0 then
						%>
							<tr class="m_l_list_b" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);">
								<td class="m_l_list_t" style="height: 26px; padding-left: 5px;">No Departments Found! Want To <a href="dept.asp?id=2">Add</a> One?</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						<%
						else

						do until RSDEPT.EOF
						%>
							<tr class="m_l_list_b" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);" onMouseOut="colorRowLight(this,0);" onmouseup="location.href='dept.asp?id=2&amp;dept=<%=RSDEPT("ID")%>'">
								<td class="m_l_list_t" style="height: 26px; padding-left: 5px;"><%=RSDEPT("FULL")%></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						<%
						RSDEPT.MOVENEXT
						loop
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
		</div>
	</div>
<%
elseif pagetype = "4" then

	if (request("gd")) = "1" then
%>
	<div class="m_l">
		<div class="m_l_title">Congratulations!</div>
		<div class="m_l_subtitle">The Department You Selected Has Been Removed!</div>
		<div class="m_l_ins">Now Please Choose An Option From Below.</div>
		<div class="m_l_sel">
			Congratulations, <%=session("sess_fn")%>! The Department You Asked To Delete Has Been Removed Successfully.<br>
			Please Choose An Option From Below To Continue.			
			<div class="botopts">
				<ul>
					<li><a href="dept.asp?id=4">Delete Another Department</a></li>
					<li><a href="dept.asp?id=2">Add A Department</a></li>
					<li><a href="dept.asp?id=1">Return To The Department Management Homepage</a></li>
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
		<div class="m_l_subtitle">The Department You Specified Cannot Be Found!</div>
		<div class="m_l_ins">Now Please Choose An Option From Below.</div>
		<div class="m_l_sel">
			Sorry, <%=session("sess_fn")%>, But The System Is Unable To Find The Department You Specified.<br>
			Please Go Back And Try Again.			
			<div class="botopts">
				<ul>
					<li><a href="#" onmouseup="history.back();">Go Back</a></li>
					<li><a href="dept.asp?id=4">Delete Another Department</a></li>
					<li><a href="dept.asp?id=2">Add A Department</a></li>
					<li><a href="dept.asp?id=1">Return To The Department Management Homepage</a></li>
					<li><a href="default.asp?id=1">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	else
%>
	<div class="m_l">
		<div class="m_l_title">Delete A Department</div>
		<div class="m_l_subtitle">Remove A Department From The System's Database.</div>
		<div class="m_l_ins">Please Select The Department You Wish To Delete.</div>
		<div class="m_l_sel">
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('added');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_dept.gif" border="0" alt="Which Department Do You Wish To Delete?"></td>
					<td class="m_l_list_t"><b>Which Department Do You Wish To Delete?</b></td>
				</tr>
				<tr id="list_added">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">

						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
						<%
						RSDEPTSQL = "SELECT * FROM Departments ORDER BY SHORT ASC"
						Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
						RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic

						if RSDEPT.RECORDCOUNT = 0 then
						%>
							<tr class="m_l_list_b" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);">
								<td class="m_l_list_t" style="height: 26px; padding-left: 5px;">No Departments Found! Want To <a href="dept.asp?id=2">Add</a> One?</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						<%
						else

						do until RSDEPT.EOF
						%>
							<tr class="m_l_list_b" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);" onmouseup="location.href='/pt/modules/ss/db/delete.asp?deltype=5&amp;dept=<%=RSDEPT("ID")%>'">
								<td class="m_l_list_t" style="height: 26px; padding-left: 5px;"><%=RSDEPT("FULL")%></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						<%
						RSDEPT.MOVENEXT
						loop
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
		</div>
	</div>
<%
	end if

elseif pagetype = "5" then

	if (request("dept")) <> "" then

		RSDEPTSQL = "SELECT * FROM Departments WHERE ID = " & request("dept")
		Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
		RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic
%>
	<div class="m_l">
		<div class="m_l_title">Actions Available</div>
		<div class="m_l_subtitle">What You Can Do To <%=RSDEPT("SHORT")%></div>
		<div class="m_l_ins">Available Actions For The Selected Department Are Listed Below.</div>
		<div class="m_l_sel">
			<div class="botopts">
				<ul>
					<li><a href="dept.asp?id=2&amp;dept=<%=RSDEPT("ID")%>">Edit The Department's Details</a></li>
					<li><a href="/pt/modules/ss/db/delete.asp?deltype=5&amp;dept=<%=RSDEPT("ID")%>">Delete The Department (No Conformation)</a><p></li>
					<li><a href="dept.asp?id=5">Return To Viewing All The Departments</a></li>
					<li><a href="dept.asp?id=1">Return To The Department Management Homepage</a></li>
					<li><a href="default.asp?id=1">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
		RSDEPT.close
		set RSDEPT = nothing

	else
%>
	<div class="m_l">
		<div class="m_l_title">View Current Departments</div>
		<div class="m_l_subtitle">Shows You All The Departments Being Used In The System.</div>
		<div class="m_l_ins">To View Any Actions You Can Do With A Department, Just Click It.</div>
		<div class="m_l_sel">
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('added');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_dept.gif" border="0" alt="Displaying All Departments"></td>
					<td class="m_l_list_t"><b>Displaying All Departments</b></td>
				</tr>
				<tr id="list_added">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">

						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
						<%
						RSDEPTSQL = "SELECT * FROM Departments ORDER BY SHORT ASC"
						Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
						RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic

						if RSDEPT.RECORDCOUNT = 0 then
						%>
							<tr class="m_l_list_b" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);">
								<td class="m_l_list_t" style="height: 26px; padding-left: 5px;">No Departments Found! Want To <a href="dept.asp?id=2">Add</a> One?</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						<%
						else

						do until RSDEPT.EOF
						%>
							<tr class="m_l_list_b" onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);" onmouseup="location.href='dept.asp?id=5&amp;dept=<%=RSDEPT("ID")%>'">
								<td class="m_l_list_t" style="height: 26px; padding-left: 5px;"><%=RSDEPT("FULL")%></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						<%
						RSDEPT.MOVENEXT
						loop
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
		</div>
	</div>
<%
	end if
else
%>
	<div class="m_l">
		<div class="m_l_title">Department Management</div>
		<div class="m_l_subtitle">Add Or Remove Departments From The System.</div>
		<div class="m_l_ins">Please Choose A Task From Below...</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_t_l"><a href="dept.asp?id=2"><img src="/pt/media/icons/48_dept.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="dept.asp?id=2">Add A Department</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="dept.asp?id=3"><img src="/pt/media/icons/48_edit.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="dept.asp?id=3">Edit A Department's Details</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="dept.asp?id=4"><img src="/pt/media/icons/48_del.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="dept.asp?id=4">Delete A Department</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="dept.asp?id=5"><img src="/pt/media/icons/48_dept.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="dept.asp?id=5">View Current Departments</a></td>
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