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
<%
if pagetype = "2" then
%>
<script language="javascript" type="text/javascript">
	function pjump2()
	{
		if (document.add.p3.value.length == 3)
			{
			document.add.p4.focus();
			}
	}
</script>
<%
else
end if
%>
<title><%=var_ptitle%></title>
</head>

<body>

<div class="smlb_b"></div>
<div class="topb_b"></div>

<div class="main">
	<!--#include virtual="/pt/modules/ss/topbar/admin.inc"-->

<%
if pagetype = "2" then
	if (request("gd") = "1") then
%>
	<div class="m_l">
		<div class="m_l_title">Congratulations!</div>
		<div class="m_l_subtitle">Your Details Have Been Successfully Changed!</div>
		<div class="m_l_ins">Your Account Details Have Been Updated With Your New Information.</div>
		<div class="m_l_sel">
			Congratulations <%=session("sess_fn")%>, Your Account Has Been Updated!<br>
			It Is Recommended That You Logout And Back In, But It Isn't Essential.<br><br>
			<b>If You Didn't Change Anything, Then Nothing Has Been Changed In Your Account!</b>
			<div class="botopts">
				<ul>
					<li><a href="/pt/modules/ss/usersys/logout.asp?id=2">Logout</a></li>
					<li><a href="default.asp?id=1">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
<%
	else
%>
	<div class="m_l">
		<div class="m_l_title">Change <i>My</i> Settings</div>
		<div class="m_l_subtitle">Edit Your Personal Details.</div>
		<div class="m_l_ins">Please Change Any Incorrect Information Below, Then Click "Save".</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<%
			if (request("err") = "1") then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t" style="color: #F00; font-weight: bold;">You Have Left One Or More Fields Blank! Please Try Again!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err") = "2") then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t" style="color: #F00; font-weight: bold;">The Two New Passwords You Entered Do Not Match! Please Try Again!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err") = "3") then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t" style="color: #F00; font-weight: bold;">Your Old Password Is Incorrect, So Your New Password Cannot Be Added!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err") = "4") then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t" style="color: #F00; font-weight: bold;">An Unknown Error Occured...</td>
				</tr>
			</table>
			<hr size="1">
			<%
			else
			end if

			if (session("sess_ttid") = "0") or (session("sess_ttid") = "") or (session("sess_ttid") = null) then
				ttid = 0
			else
				ttid = session("sess_ttid")
			end if

			RSUSERSQL = "SELECT * FROM Timetables WHERE ID = " & ttid
			RSACCSQL = "SELECT * FROM Admin WHERE UN = '" & session("sess_un") & "'"
			RSDEPTSQL = "SELECT * FROM Departments ORDER BY SHORT"
			RSROOMSQL = "SELECT * FROM Rooms ORDER BY ROOMNO ASC"

			Set RSUSER = Server.CreateObject("Adodb.RecordSet")
			RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic

			Set RSACC = Server.CreateObject("Adodb.RecordSet")
			RSACC.Open RSACCSQL, userconn, adopenkeyset, adlockoptimistic
				
			Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
			RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic

			Set RSROOM = Server.CreateObject("Adodb.RecordSet")
			RSROOM.Open RSROOMSQL, dataconn, adopenkeyset, adlockoptimistic

			if RSUSER.RECORDCOUNT = 0 then
				if RSACC.RECORDCOUNT = 0 then
			%>
	<div class="m_l">
		<div class="m_l_title">Sorry!</div>
		<div class="m_l_subtitle">Your Account Cannot Be Found!</div>
		<div class="m_l_ins">A Serious Error Has Occured.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			Sorry, <%=session("sess_fn")%>, But The System Cannot Find Your Details.<br>
			Please Contact The System Administrator For Help.
			<div class="botopts">
				<ul>
					<li><a href="default.asp">Return To The Admin Homepage</a></li>
				</ul>
			</div>
		</div>
	</div>
			<%
				else
					var_title = RSACC("TITLE")
					var_fn = RSACC("FN")
					var_ln = RSACC("LN")
					var_dept = "NA"
					var_pos = "NA"
					var_room = "NA"
					var_email = RSACC("EMAIL")
				end if
			else
				var_title = RSUSER("TITLE")
				var_fn = RSUSER("FN")
				var_ln = RSUSER("LN")
				var_dept = RSUSER("DEPT")
				var_pos = RSUSER("CATEGORY")
				var_room = RSUSER("DEFROOM")
				var_email = RSACC("EMAIL")
			end if
			%>
			Any Details You Change Here Will Also Update Your Name, Title Or Other Credential On Your Timetable Information.
			<b>You Cannot Change Your Username.</b>
			<form name="add" action="/pt/modules/ss/db/edit.asp?edittype=13" method="post">
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('your');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="A Little Image To Represent Your Lovely Self."></td>
					<td class="m_l_list_t"><b>Your Personal Information</b></td>
				</tr>
				<tr id="list_your">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="line-height: 20px; padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Title</b></td>
								<td>
								<select class="b_std" size="1" name="TITLE">
									<option<%if var_title = "Mr" then%> selected<%else%><%end if%>>Mr</option>
									<option<%if var_title = "Mrs" then%> selected<%else%><%end if%>>Mrs</option>
									<option<%if var_title = "Miss" then%> selected<%else%><%end if%>>Miss</option>
									<option<%if var_title = "Ms" then%> selected<%else%><%end if%>>Ms</option>
									<option<%if var_title = "Mdme" then%> selected<%else%><%end if%>>Mdme</option>
									<option<%if var_title = "Dr" then%> selected<%else%><%end if%>>Dr</option>
								</select>								
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 150px; padding-left: 5px;"><b>First Name</b></td>
								<td><input class="b_std" type="text" name="FN" size="20" value="<%=var_fn%>"></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Last Name</b></td>
								<td><input class="b_std" type="text" name="LN" size="20" value="<%=var_ln%>"></td>
							</tr>
							<%
							if var_dept = "NA" then
							else
							%>
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
									<option value="<%=RSDEPT("DEPTID")%>"<%if RSDEPT("DEPTID") = var_dept then%> selected<%else%><%end if%>><%=RSDEPT("Full")%></option>
									<%
									RSDEPT.MOVENEXT
									loop
									%>
								</select>
								</td>
							</tr>
							<%
							end if

							if var_pos = "NA" then
							else
							%>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Position</b></td>
								<td>
								<select class="b_std" size="1" name="CATEGORY">
									<option value="T"<%if var_pos = "T" then%> selected<%else%><%end if%>>Teacher</option>
									<option value="PT"<%if var_pos = "PT" then%> selected<%else%><%end if%>>Principal Teacher</option>
									<option value="DHT"<%if var_pos = "DHT" then%> selected<%else%><%end if%>>Deputy Head Teacher</option>
									<option value="HT"<%if var_pos = "HT" then%> selected<%else%><%end if%>>Head Teacher</option>
									<option value="PS"<%if var_pos = "PS" then%> selected<%else%><%end if%>>Pupil Support</option>
									<option value="OC"<%if var_pos = "OC" then%> selected<%else%><%end if%>>Outside Cover</option>
								</select>
								</td>
							</tr>
							<%
							end if

							if var_room = "NA" then
							else
							%>
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
									<option<%if var_room = RSROOM("ROOMNO") then%> selected<%else%><%end if%>><%=RSROOM("ROOMNO")%></option>
									<%
									RSROOM.MOVENEXT
									loop
									%>
								</select>
								</td>
							</tr>
							<%
							end if
							%>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>E-Mail</b></td>
								<td style="font-size: 10pt;">
								<input class="b_std" type="text" name="email" size="20" value="<%=removedomain(var_email)%>">&nbsp;@<%=var_emaildomain1%>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('acc');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="A Little Image To Represent Your Lovely Self."></td>
					<td class="m_l_list_t"><b>Your PleaseTakes Account Details</b></td>
				</tr>
				<tr id="list_acc">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="line-height: 20px; padding-top: 7px;">
					If You <b>Don't</b> Want To Change Your Password<%if var_est_enabled_pin = "1" then%> Or PIN<%else%><%end if%>, Just Leave These Boxes Blank.
					Your Old Password Is Required For Security. If You Leave The "Old Password" Field Blank And Enter A New Password In The Other Two Fields, The System Will <b>Not</b> Update It.
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 160px; padding-left: 5px;"><b>Old Password</b></td>
								<td style="font-size: 10pt;">
								<input class="b_pwd" type="password" name="password_o" size="20">
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>New Password</b></td>
								<td style="font-size: 10pt;">
								<input class="b_pwd" type="password" name="password_n" size="20">
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Confirm New Password</b></td>
								<td style="font-size: 10pt;">
								<input class="b_pwd" type="password" name="password_c" size="20">
								</td>
							</tr>
							<%
							if var_est_enabled_pin = "1" then
							%>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Old PIN</b></td>
								<td style="font-size: 10pt;">
								<input class="b_pwd" type="password" name="p1" size="3" maxlength="3" onkeyup="pjump();">&nbsp;/&nbsp;<input class="b_pwd" type="password" name="p2" size="3" maxlength="3">
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>New PIN</b></td>
								<td style="font-size: 10pt;">
								<input class="b_pwd" type="password" name="p3" size="3" maxlength="3" onkeyup="pjump2();">&nbsp;/&nbsp;<input class="b_pwd" type="password" name="p4" size="3" maxlength="3">
								</td>
							</tr>
							<%
							else
							end if
							%>
						</table>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2">
					<hr size="1">
					<div style="width: 100%; font-family: Tahoma,Sans-Serif; font-size: 10pt; text-align: center;"><b><a href="#" onmouseup="document.add.submit();">Save</a></b> :: <b><a href="#" onmouseup="document.add.reset();">Reset</a></b></div>
					</td>
				</tr>
			</table>
			</form>
		<%
		RSDEPT.close
		RSROOM.close
		set RSDEPT = nothing
		set RSROOM = nothing
		%>
		</div>
	</div>
<%
	end if
elseif pagetype = "3" then
%>
<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->
	<form name="edit" action="/pt/modules/ss/db/edit.asp?edittype=6" method="post">
	<div class="m_l">
		<div class="m_l_title">Enable/Disable Features</div>
		<div class="m_l_subtitle">Enable Or Disable Features On The System.</div>
		<div class="m_l_ins">Please Change Any Required Settings Below.</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">
					<%
					if (request("ok")) = "1" then
					%>
					<span style="color: #389F0A;"><b>Settings Successfully Changed</b></span>
					<%
					else
					%>
					If The Settings Are Changed Successfully, A Notification In Green Will Appear Here.
					<%
					end if
					%>
					</td>
				</tr>
			</table>
			<hr size="1">
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('status');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_cog.gif" border="0" alt="Question Mark"></td>
					<td class="m_l_list_t"><b>Entire System Status</b></td>
				</tr>
				<tr id="list_status">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr>
								<td class="m_l_list_t" style="padding-left: 5px;">
								<b>Warning</b> - Disabling This System Prevents Standard Users And Standard Admin Accounts From Logging In!
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
							</tr>
							<tr>
							<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
								<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
									<td class="m_l_list_t" style="height: 28px; padding-left: 5px;" onmouseup="document.edit.SYSTEM[0].checked = true; document.edit.submit();"><input type="radio" value="1" name="SYSTEM" id="ENABLED"<%if var_est_enabled = "1" then%> checked<%else%><%end if%>><label for="ENABLED">&nbsp;<b>Enabled</b></label></td>
								</tr>
								<tr>
									<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
								</tr>
								<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
									<td class="m_l_list_t" style="height: 28px; padding-left: 5px;" onmouseup="document.edit.SYSTEM[1].checked = true; document.edit.submit();"><label for="DISABLED"><input type="radio" value="0" name="SYSTEM" id="DISABLED"<%if var_est_enabled = "0" then%> checked<%else%><%end if%>>&nbsp;<b>Disabled</b></label></td>
								</tr>
							</table>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep"></td>
				</tr>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('signup');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_cog.gif" border="0" alt="Question Mark"></td>
					<td class="m_l_list_t"><b>User Signups</b></td>
				</tr>
				<tr id="list_signup">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr>
								<td class="m_l_list_t" style="padding-left: 5px;">
								Disabling This Feature Will Prevent New Users From Signing Up From The Login Screen.
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;" onmouseup="document.edit.SIGNUP[0].checked = true; document.edit.submit();"><input type="radio" value="1" name="SIGNUP" id="SIGNUP_ENABLED"<%if var_est_enabled_signup = "1" then%> checked<%else%><%end if%>><label for="SIGNUP_ENABLED">&nbsp;<b>Enabled</b></label></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;" onmouseup="document.edit.SIGNUP[1].checked = true; document.edit.submit();"><label for="SIGNUP_DISABLED"><input type="radio" value="0" name="SIGNUP" id="SIGNUP_DISABLED"<%if var_est_enabled_signup = "0" then%> checked<%else%><%end if%>>&nbsp;<b>Disabled</b></label></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep"></td>
				</tr>
			</table>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('PIN');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_cog.gif" border="0" alt="Question Mark"></td>
					<td class="m_l_list_t"><b>Using PINs</b></td>
				</tr>
				<tr id="list_PIN">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr>
								<td class="m_l_list_t" style="padding-left: 5px;">
								A Six-Digit PIN Number Can Also Be Used At The Login Screen For Extra Security.<br>
								If This Will Be The First Time You Enable This Feature, Everyone's PIN Will Be Blank, And They Will
								Receive An Alert, Asking Them To Add One.
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;" onmouseup="document.edit.PIN[0].checked = true; document.edit.submit();"><input type="radio" value="1" name="PIN" id="PIN_ENABLED"<%if var_est_enabled_pin = "1" then%> checked<%else%><%end if%>><label for="PIN_ENABLED">&nbsp;<b>Enabled</b></label></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;" onmouseup="document.edit.PIN[1].checked = true; document.edit.submit();"><label for="PIN_DISABLED"><input type="radio" value="0" name="PIN" id="PIN_DISABLED"<%if var_est_enabled_pin = "0" then%> checked<%else%><%end if%>>&nbsp;<b>Disabled</b></label></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</div>
	</div>
	</form>
<%
elseif pagetype = "4" then
%>
<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->
	<form name="edit" action="/pt/modules/ss/db/edit.asp?edittype=4" method="post">
	<div class="m_l">
		<div class="m_l_title">Period/Day Settings</div>
		<div class="m_l_subtitle">Change Whether The System Uses Weekends And Day Settings.</div>
		<div class="m_l_ins">Please Change Any Settings Below.</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">
					<%
					if (request("ok")) = "1" then
					%>
					<span style="color: #389F0A;"><b>Settings Successfully Changed</b></span>
					<%
					else
					%>
					If The Settings Are Changed Successfully, A Notification In Green Will Appear Here.
					<%
					end if
					%>
					</td>
				</tr>
			</table>
			<hr size="1">
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('weekends');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_cog.gif" border="0" alt="Does <%=var_est_full%> Require Weekends?"></td>
					<td class="m_l_list_t"><b>Weekend Capability</b></td>
				</tr>
				<tr id="list_weekends">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr>
								<td class="m_l_list_t" style="padding-left: 5px;">Does <%=var_est_full%> Require The Use Of Weekends? Selecing "Yes" Will Enable The Use Of Saturdays And Sundays, And Will Ensure Data Backup Will Occur On A Saturday.</td>
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
						</table>
					</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
			</table>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('periods'); showdetail('tt');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_cog.gif" border="0" alt="How Many Periods In Each Day?"></td>
					<td class="m_l_list_t"><b>Periods In Each Day</b></td>
				</tr>
				<tr id="list_periods">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr>
								<td class="m_l_list_t" style="padding-left: 5px;">The Timetable Below Shows You How Many Periods Each Day Has.<br>
								The Maximum Number Of Periods That Can Be Used Is <b><%=var_maxperiods%></b>.<br>
								<span style="font-weight: bold; color: #498F32;">Green</span> Periods Mean That There Are Classes Running For The Day In Question.<br>
								<span style="font-weight: bold; color: #F00;">Red</span> Simply Means There Isn't.
								Simply Click A <span style="font-weight: bold; color: #F00;">Red</span> Or <span style="font-weight: bold; color: #498F32;">Green</span> Period To Enlarge Or Reduce The Number Of Periods In Each Day.
								If All The Days Have The Same Number Of Periods, Just Click The Number On The Top Row, And All Days Will Be Updated.
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div id="list_tt">
			<%
			if var_est_enabled_weekends = "1" then
			%>
			<!--#include virtual="/pt/modules/ss/timetables/admin_set_7day.inc"-->
			<%
			else
			%>
			<!--#include virtual="/pt/modules/ss/timetables/admin_set_5day.inc"-->
			<%
			end if
			%>
			</div>
			<div style="height: 10px;"></div>
		</div>
	</div>
	</form>
<%
elseif pagetype = "5" then
%>
<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->
	<form name="edit" action="/pt/modules/ss/db/edit.asp?edittype=7" method="post">
	<div class="m_l">
		<div class="m_l_title">Change Establishment Information</div>
		<div class="m_l_subtitle">Change Information On The Establishment Using The System.</div>
		<div class="m_l_ins">Please Change Any Incorrect Information Below, Then Click "Save".</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">
					<%
					if (request("ok")) = "1" then
					%>
					<span style="color: #389F0A;"><b>Settings Successfully Changed</b></span>
					<%
					else
					%>
					If The Settings Are Changed Successfully, A Notification In Green Will Appear Here.
					<%
					end if
					%>
					</td>
				</tr>
			</table>
			<hr size="1">
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('estnames');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_cog.gif" border="0" alt="Question Mark"></td>
					<td class="m_l_list_t"><b>Names Of Your Establishment</td>
				</tr>
				<tr id="list_estnames">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr>
								<td class="m_l_list_t" style="padding-left: 5px; line-height: 20px;" colspan="2">
								The List Below Shows You Information Held About <%=var_est_full%>. If Any Of These Details Are
								Incorrect, Please Change Them And Click "Save".
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
							</tr>
						</table>
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="width: 150px; height: 28px; padding-left: 5px;"><b>Full Name</b></td>
								<td>
								<input class="b_std" type="text" name="FULL" size="40" value="<%=var_est_full%>">
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="width: 150px; height: 28px; padding-left: 5px;"><b>Short Name (Initials)</b></td>
								<td>
								<input class="b_std" type="text" name="SHORT" size="40" value="<%=var_est_short%>">
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						</table>
					<hr size="1">
					<div style="width: 100%; text-align: center;"><b><a href="#" onmouseup="document.edit.submit();">Save</a></b> :: <b><a href="#" onmouseup="document.edit.reset();">Reset</a></b></div>
					</td>
				</tr>
			</table>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('staff');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_cog.gif" border="0" alt="Question Mark"></td>
					<td class="m_l_list_t"><b>Your Staff Titles</td>
				</tr>
				<tr id="list_staff">
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr>
								<td class="m_l_list_t" style="padding-left: 5px; line-height: 20px;" colspan="2">
								The Fields Below Determine What Staff Are Called, Depending On Which Section Of The System They Login To.
								These Names Are Used On The Login Screens.
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" style="height: 5px;" colspan="2"></td>
							</tr>
						</table>
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="width: 150px; height: 28px; padding-left: 5px;"><b>Standard Users</b></td>
								<td>
								<input class="b_std" type="text" name="STD" size="40" value="<%=var_usernames_std%>">
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="width: 150px; height: 28px; padding-left: 5px;"><b>Administrators</b></td>
								<td>
								<input class="b_std" type="text" name="ADMIN" size="40" value="<%=var_usernames_admin%>">
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						</table>
					<hr size="1">
					<div style="width: 100%; text-align: center;"><b><a href="#" onmouseup="document.edit.submit();">Save</a></b> :: <b><a href="#" onmouseup="document.edit.reset();">Reset</a></b></div>
					</td>
				</tr>
			</table>
		</div>
	</div>
	</form>
<%
else
%>
<!--#include virtual="/pt/modules/ss/usersys/admincheck_1.inc"-->
	<div class="m_l">
		<div class="m_l_title">Change Settings</div>
		<div class="m_l_subtitle">Change Settings Vital To The System.</div>
		<div class="m_l_ins">Please Choose A Task From Below...</div>
		<div class="m_l_sel">		
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_t_l"><a href="settings.asp?id=3"><img src="/pt/media/icons/48_endis.png" border="0" alt="Enable Or Disable Features On The System."></a></td>
					<td class="m_l_sel_t_r"><a href="settings.asp?id=3">Enable/Disable Features</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="settings.asp?id=4"><img src="/pt/media/icons/48_clock.png" border="0" alt="Change Whether Or Not That The System Will Use Weekends, Or Specify How Many Periods A Day Has."></a></td>
					<td class="m_l_sel_t_r"><a href="settings.asp?id=4">Change Period/Day Settings</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="settings.asp?id=5"><img src="/pt/media/icons/48_est.png" border="0" alt="Click Here To Change Information On The Establishment Using The System."></a></td>
					<td class="m_l_sel_t_r"><a href="settings.asp?id=5">Change Establishment Information</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
			</table>
			If You Require Use Of 11 Or More Periods A Day, Please Get In Contact With Server-ML, Who Will Be Happy To Add The Number Of Periods You Require.
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