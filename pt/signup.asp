<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" >

<!--#include virtual="/pt/modules/ss/p_s.inc"-->
<%
pagetype = request("id")
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="/pt/modules/css/std.css">
<script language="javascript" type="text/javascript" src="/pt/modules/js/signup.js"></script>
<title><%=var_ptitle%></title>
</head>

<body<%if pagetype = "4" then%> onload="document.frm.UN.focus();"<%else%><%end if%>>

<div class="smlb_b"></div>
<div class="topb_b"></div>

<div class="main" style="width: 634px;">
	<!--#include virtual="/pt/modules/ss/topbar/signup_popup.inc"-->
<%
if pagetype = "2" then
%>
	<div class="topb_m">
		<ul>
			<li><b>Step 1</b> :: </li>
			<li>Step 2 :: </li>
			<li>Step 3 :: </li>
			<li>Complete :: </li>
			<li><a href="#" onmouseup="self.close();">Close Popup</a></li>
		</ul>
	</div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Step 1</div>
		<div class="m_l_subtitle">Select Your Department</div>
		<div class="m_l_ins">Please Choose Your Department From Below.</div>
		<div class="m_l_sel">
			All The Departments Are Listed Below.<br>
			Please Select The One You Belong To, And You Will Be Taken To Step 2!
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
				RSDEPTSQL = "SELECT * FROM Departments ORDER BY Short"
					
				Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
				RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic
				
				do until RSDEPT.EOF
				%>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="location.href='signup.asp?id=3&amp;dept=<%=RSDEPT("DEPTID")%>'">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_dept.gif" border="0" alt="The <%=RSDEPT("FULL")%> Department"></td>
					<td class="m_l_list_t"><b><%=RSDEPT("FULL")%></b></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
				RSDEPT.MOVENEXT
				loop
				
				RSDEPT.close
				set RSDEPT = nothing
				%>
			</table>
		</div>
	</div>
<%
elseif pagetype = "3" then

	if (request("DEPT")) = "" then
		response.redirect "/pt/signup.asp?id=2"
	else
	end if
%>
	<div class="topb_m">
		<ul>
			<li><a href="signup.asp?id=2">Step 1</a> :: </li>
			<li><b>Step 2</b> :: </li>
			<li>Step 3 :: </li>
			<li>Complete :: </li>
			<li><a href="#" onmouseup="self.close();">Close Popup</a></li>
		</ul>
	</div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Step 2</div>
		<div class="m_l_subtitle">Select Yourself!</div>
		<div class="m_l_ins">Please Select Your Name From The List Below.</div>
		<div class="m_l_sel">
			If Your Name Isn't On The List, Please Click <a href="#"><b>Here</b></a>.
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
				RSUSERSQL = "SELECT * FROM Timetables WHERE DEPT = " & request("DEPT") & " ORDER BY LN"
					
				Set RSUSER = Server.CreateObject("Adodb.RecordSet")
				RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
				
				do until RSUSER.EOF

					RSCHECKSQL = "SELECT * FROM Users WHERE TTID = " & RSUSER("ID")
					
					Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
					RSCHECK.Open RSCHECKSQL, userconn, adopenkeyset, adlockoptimistic
					
					RSCHECK2SQL = "SELECT * FROM Admin WHERE TTID = " & RSUSER("ID")

					Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
					RSCHECK2.Open RSCHECK2SQL, userconn, adopenkeyset, adlockoptimistic

					if (RSCHECK.RECORDCOUNT => 1) OR (RSCHECK2.RECORDCOUNT => 1) then
						RSCHECK.close
						set RSCHECK = nothing
						RSCHECK2.close
						set RSCHECK2 = nothing
					else
				%>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="location.href='signup.asp?id=4&amp;uid=<%=RSUSER("ID")%>&amp;dept=<%=request("dept")%>'">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="<%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%>"></td>
					<td class="m_l_list_t"><b><%=RSUSER("FN")%>&nbsp;<%=RSUSER("LN")%></b></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<%
						RSCHECK.close
						set RSCHECK = nothing
						RSCHECK2.close
						set RSCHECK2 = nothing
					end if
				
				RSUSER.MOVENEXT
				loop
				
				RSUSER.close
				set RSUSER = nothing
				%>
			</table>
		</div>
	</div>
<%
elseif pagetype = "4" then
%>
	<div class="topb_m">
		<ul>
			<li><a href="signup.asp?id=2">Step 1</a> :: </li>
			<li><a href="signup.asp?id=3&dept=<%=request("dept")%>">Step 2</a> :: </li>
			<li><b>Step 3</b> :: </li>
			<li>Complete :: </li>
			<li><a href="#" onmouseup="self.close();">Close Popup</a></li>
		</ul>
	</div>
	<form name="frm" action="/pt/modules/ss/db/add.asp?addtype=5&amp;uid=<%=request("uid")%>&amp;dept=<%=request("dept")%>" method="post">
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Step 3</div>
		<div class="m_l_subtitle">The Final Step!</div>
		<div class="m_l_ins">Please Fill Out ALL The Blanks Below, And Click "Create Account".</div>
		<div class="m_l_sel">
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
					<td class="m_l_list_t">The Two Passwords You Entered Don't Match! Please Try Again!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err")) = "4" then
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
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="That's You!"></td>
					<td class="m_l_list_t"><b>Your New Account Details</b></td>
				</tr>
				<tr>
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 150px; padding-left: 5px;"><b>Username</b></td>
								<td><input class="b_std" type="text" name="UN" size="20"></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 150px; padding-left: 5px;"><b>Password</b></td>
								<td><input class="b_pwd" type="password" name="PW" size="20"></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 150px; padding-left: 5px;"><b>Verify Password</b></td>
								<td><input class="b_pwd" type="password" name="PWV" size="20"></td>
							</tr>
							<%
							if var_est_enabled_pin = "1" then
							%>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 150px; padding-left: 5px;"><b>PIN</b></td>
								<td style="font-size: 10pt;"><input class="b_pwd" type="password" name="p1" size="3" maxlength="3" onkeyup="pjump();">&nbsp;/&nbsp;<input class="b_pwd" type="password" name="p2" size="3" maxlength="3"></td>
							</tr>
							<%
							else
							end if
							%>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 150px; padding-left: 5px;"><b>E-Mail</b></td>
								<td style="font-size: 10pt;"><input class="b_std" type="text" name="EMAIL" size="20">&nbsp;@<%=var_emaildomain1%></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
						</table>
					Your E-Mail Address Is Required For A Feature That Will Be Able To Send You Notifications About A PleaseTake Via E-Mail.
					<hr size="1">
					<div style="width: 100%; text-align: center;"><b><a href="#" onmouseup="document.frm.submit();">Create Account</a></b> :: <b><a href="#" onmouseup="document.frm.reset();">Clear</a></b></div>
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
elseif pagetype = "5" then
%>
	<div class="topb_m">
		<ul>
			<li>Step 1 :: </li>
			<li>Step 2 :: </li>
			<li>Step 3 :: </li>
			<li><b>Complete</b> :: </li>
			<li><a href="#" onmouseup="self.close();">Close Popup</a></li>
		</ul>
	</div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Congratulations!</div>
		<div class="m_l_subtitle">You Have Signed Up!</div>
		<div class="m_l_ins">All The Details Have Been Successfully Saved, And You Can Now Login.</div>
		<div class="m_l_sel">
			Congratulations, You Now Have An Account On The System!<br>
			Everything You Entered Has Been Successfully Saved, And All You Need To Do Now Is Close This Popup
			Window By Clicking "Finish" Below, And Logging In!			
			<hr size="1">
			<div style="width: 634px; text-align: right; font-size: 18pt; font-weight: bold;"><a href="#" onmouseup="opener.location.reload(true); self.close();">Finish</a></div>
		</div>
	</div>
<%
else
%>
	<div class="topb_m">
		<ul>
			<li><a href="signup.asp?id=2">Step 1</a> :: </li>
			<li>Step 2 :: </li>
			<li>Step 3 :: </li>
			<li>Complete :: </li>
			<li><a href="#" onmouseup="self.close();">Close Popup</a></li>
		</ul>
	</div>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title"><script language="javascript" type="text/javascript">document.write(daymsg)</script>!</div>
		<div class="m_l_subtitle">Welcome To The System...</div>
		<div class="m_l_ins">...And Welcome To The Signup Wizard!</div>
		<div class="m_l_sel">
			Hello, And Thank You For Deciding To Join The System!<br>
			The Signup Process Is Very Simple, Just Choose Yourself From A List, Type In A Username And Password, And You're
			Away!<br><br>
			<b>It's That Simple!</b><br>
			To Begin, Just Click "Next" Below.
			<hr size="1">
			<div style="width: 634px; text-align: right; font-size: 18pt; font-weight: bold;"><a href="signup.asp?id=2">Next></a></div>
		</div>
	</div>
<%
end if
%>
</div>

</body>

</html>

<!--#include virtual="/pt/modules/ss/p_e.inc"-->