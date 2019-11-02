<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" >
<!--#include virtual="/pt/modules/ss/usersys/logincheck_admin.inc"-->

<!--#include virtual="/pt/modules/ss/p_s.inc"-->
<%
pagetype = request("id")

if session("sess_un") = settingsXML.documentElement.childNodes.item(0).childNodes.item(4).getAttribute("firstloginacc") then
	response.redirect "setup.asp?err=2"
else
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
	if isNumeric(request("uid")) = False then
		response.redirect "/pt/admin/ocover.asp?id=1"
	elseif request("uid") = "" then
		response.redirect "/pt/admin/ocover.asp?id=1"
	else
		RSWEEKSQL = "SELECT * FROM PERIODS"
		Set RSWEEK = Server.CreateObject("Adodb.RecordSet")
		RSWEEK.Open RSWEEKSQL, dataconn, adopenkeyset, adlockoptimistic
	
		if var_est_enabled_weekends = "1" then
			for i = 1 to 7
			weektotal = weektotal + RSWEEK("TOTALS")
			RSWEEK.MOVENEXT
			next
		else
			RSWEEK.MOVENEXT
			for i = 2 to 6
			weektotal = weektotal + RSWEEK("TOTALS")
			RSWEEK.MOVENEXT
			next
		end if	

		RSCOVERSQL = "SELECT * FROM OCover WHERE ID = " & request("uid")
		Set RSCOVER = Server.CreateObject("Adodb.RecordSet")
		RSCOVER.Open RSCOVERSQL, dataconn, adopenkeyset, adlockoptimistic
		
		if RSCOVER.RECORDCOUNT = 0 then
			response.redirect "/pt/admin/ocover.asp?id=1"
		else
%>
	<div class="m_l">
		<div class="m_l_title">Outside Cover</div>
		<div class="m_l_subtitle">In Trouble? Get Some Help From The Outside Here.</div>
		<div class="m_l_ins">Please Choose A Task From Below...</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<%
			if (request("err") = "1") then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Left Some Details Out In The Form! Please Try Again!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			end if
			%>
			Here, You Can Change Details For <%=RSCOVER("FN")%>&nbsp;<%=RSCOVER("LN")%>, Such His/Her Name, Title Or Entitlement.
			Once You Are Done, Click "Save".
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td colspan="2" style="height: 10px;"></td>
				</tr>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="<%=RSCOVER("FN")%>&nbsp;<%=RSCOVER("LN")%>"></td>
					<td class="m_l_list_t"><b><%=RSCOVER("FN")%>&nbsp;<%=RSCOVER("LN")%>'s Details</b></td>
				</tr>
				<tr>
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<form name="edit" action="/pt/modules/ss/db/edit.asp?edittype=14&amp;uid=<%=RSCOVER("ID")%>" method="post">
						<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Title</b></td>
								<td>
								<select class="b_std" size="1" name="TITLE">
									<option<%if RSCOVER("TITLE") = "Mr" then%> selected<%else%><%end if%>>Mr</option>
									<option<%if RSCOVER("TITLE") = "Mrs" then%> selected<%else%><%end if%>>Mrs</option>
									<option<%if RSCOVER("TITLE") = "Miss" then%> selected<%else%><%end if%>>Miss</option>
									<option<%if RSCOVER("TITLE") = "Ms" then%> selected<%else%><%end if%>>Ms</option>
									<option<%if RSCOVER("TITLE") = "Mdme" then%> selected<%else%><%end if%>>Mdme</option>
									<option<%if RSCOVER("TITLE") = "Dr" then%> selected<%else%><%end if%>>Dr</option>
								</select>								
								</td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; width: 150px; padding-left: 5px;"><b>First Name</b></td>
								<td><input class="b_std" type="text" name="FN" size="20" value="<%=RSCOVER("FN")%>"></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Last Name</b></td>
								<td><input class="b_std" type="text" name="LN" size="20" value="<%=RSCOVER("LN")%>"></td>
							</tr>
							<tr>
								<td class="m_l_sel_sep" colspan="2"></td>
							</tr>
							<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRow(this,0);">
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Period Entitlement</b></td>
								<td style="font-size: 10pt;">
								<input class="b_std" type="text" name="ENTITLEMENT" size="2" maxlength="2" value="<%=RSCOVER("ENTITLEMENT")%>" onblur="noalpha(this.value);"> Periods/Week (Current Week Total Is <b><%=weektotal%></b>)
								</td>
							</tr>
						</table>
						<hr size="1">
						<div style="width: 100%; text-align: center;"><b><a href="#" onmouseup="document.edit.submit();">Save</a></b> :: <b><a href="#" onmouseup="document.edit.reset();">Reset</a></b></div>
						</form>
					</td>
				</tr>
			</table>
		</div>
	</div>
<%
		end if

		RSCOVER.close
		set RSCOVER = nothing
		
		RSWEEK.close
		set RSWEEK = nothing
	end if
else
	RSWEEKSQL = "SELECT * FROM PERIODS"
	Set RSWEEK = Server.CreateObject("Adodb.RecordSet")
	RSWEEK.Open RSWEEKSQL, dataconn, adopenkeyset, adlockoptimistic
	
	if var_est_enabled_weekends = "1" then
		for i = 1 to 7
			weektotal = weektotal + RSWEEK("TOTALS")
			RSWEEK.MOVENEXT
		next
	else
		RSWEEK.MOVENEXT
		for i = 2 to 6
		weektotal = weektotal + RSWEEK("TOTALS")
		RSWEEK.MOVENEXT
		next
	end if	
%>
	<div class="m_l">
		<div class="m_l_title">Outside Cover</div>
		<div class="m_l_subtitle">In Trouble? Get Some Help From The Outside Here.</div>
		<div class="m_l_ins">Please Choose A Task From Below...</div>
		<div class="m_l_sel">
			<!--#include virtual="/pt/modules/ss/alerts/admin_alerts.inc"-->
			<%
			if (request("err") = "1") then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Left Some Details Out In The Form! Please Try Again!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err") = "2") then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">The Member You Entered Already Exists!</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("err") = "3") then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">Your Browser Passed An Invalid String To The Server. Please Try Again.</td>
				</tr>
			</table>
			<hr size="1">
			<%
			elseif (request("ok") = "1") then
			%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t"><b>Details Successfully Updated!</b></td>
				</tr>
			</table>
			<hr size="1">
			<%
			end if
			%>
			This Feature Is Used When There Aren't Enough Teaching Staff Available To Cover Every Class.
			You Can Call In Additional Staff To Help Cover The Remaining Classes.
			When You Select Someone To Be Here For Either Today Or The Whole Week, Their Name Will Be Available In The <a href="cover.asp?id=1">Staff Cover</a> Wizard.
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep" colspan="4"></td>
				</tr>
				<%
				RSLISTSQL = "SELECT * FROM OCover ORDER BY LN"
				Set RSLIST = Server.CreateObject("Adodb.RecordSet")
				RSLIST.Open RSLISTSQL, dataconn, adopenkeyset, adlockoptimistic
				
				if RSLIST.RECORDCOUNT = 0 then
				%>
				<br>
				<b>No Outside Cover Staff Were Found! Please Add Some From Below To Continue!</b>
				<%
				else
				
				currfield = "D_" & dow
				%>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_t rep_week_t" style="width: 35px; text-align: center;"></td>
					<td class="m_l_list_t rep_week_t" style="text-align: left;"><a title="Lists The Staff Names."><b>Staff Member</b></a></td>
					<td class="m_l_list_t rep_week_t"><a title="This Column Determines If The Outside Cover Member Has Been Select To Be Available To Cover For Just Today. If This Is Selected As No And The Week Column As Yes, This Means That The Staff Member Will Not Cover Today, But The Rest Of The Week."><b>Today</b></a></td>
					<td class="m_l_list_t rep_week_t"<a title="This Column Determines If The Outside Cover Member Has Been Selected To Be Available For The Entire Week."><b>Week</b></a></td>
					<td class="m_l_list_t rep_week_t" ><a title="This Column Tells You How Many Periods The Outside Cover Member Has Available To Cover Each Week."><b>Total</b></a></td>
					<td class="m_l_list_t rep_week_t" ><a title="This Column Tells You How Many Periods The Outside Cover Member Has Left For The Week. The Higher The Number, The Less They Have Covered."><b>Free</b></a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" style="height: 5px;" colspan="5"></td>
				</tr>
				<%
				do until RSLIST.EOF

				RSCOUNTSQL = "SELECT * FROM Cover WHERE COVERING = " & RSLIST("ID") & " AND OCOVER = 1"
				Set RSCOUNT = Server.CreateObject("Adodb.RecordSet")
				RSCOUNT.Open RSCOUNTSQL, dataconn, adopenkeyset, adlockoptimistic

					if (RSLIST("ENTIRE") = "1") then
				%>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);">
					<td class="m_l_list_t rep_week_t" style="width: 55px; text-align: center;"><a href="ocover.asp?id=2&amp;uid=<%=RSLIST("ID")%>"><img src="/pt/media/icons/16_edit.gif" border="0" alt="Click Here To Edit <%=RSLIST("FN")%>&nbsp;<%=RSLIST("LN")%>'s Details."></a>&nbsp;<a href="/pt/modules/ss/db/delete.asp?deltype=6&amp;uid=<%=RSLIST("ID")%>"><img src="/pt/media/icons/16_del.gif" border="0" alt="Click Here To Delete <%=RSLIST("FN")%>&nbsp;<%=RSLIST("LN")%> From The System."></a></td>
					<td class="m_l_list_t rep_week_t" style="text-align: left;"><%=RSLIST("TITLE")%>. <%=RSLIST("FN")%>&nbsp;<%=RSLIST("LN")%></td>
					<td class="m_l_list_t rep_week_t" style="width: 55px;"><a href="/pt/modules/ss/db/edit.asp?edittype=11&amp;uid=<%=RSLIST("ID")%>&amp;type=1a"><b>Yes</b></a></td>
					<td class="m_l_list_t rep_week_t" style="width: 55px;"><a href="/pt/modules/ss/db/edit.asp?edittype=11&amp;uid=<%=RSLIST("ID")%>&amp;type=1b"><b>Yes</b></a></td>
					<td class="m_l_list_t rep_week_t" style="width: 45px;"><%=RSLIST("ENTITLEMENT")%></td>
					<td class="m_l_list_t rep_week_t" style="width: 45px;"><%=RSLIST("ENTITLEMENT") - RSCOUNT.RECORDCOUNT%></td>
				</tr>
				<%
					elseif (RSLIST(currfield) = "1") and (RSLIST("ENTIRE") <> "1") then
				%>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);">
					<td class="m_l_list_t rep_week_t" style="width: 55px; text-align: center;"><a href="ocover.asp?id=2&amp;uid=<%=RSLIST("ID")%>"><img src="/pt/media/icons/16_edit.gif" border="0" alt="Click Here To Edit <%=RSLIST("FN")%>&nbsp;<%=RSLIST("LN")%>'s Details."></a>&nbsp;<a href="/pt/modules/ss/db/delete.asp?deltype=6&amp;uid=<%=RSLIST("ID")%>"><img src="/pt/media/icons/16_del.gif" border="0" alt="Click Here To Delete <%=RSLIST("FN")%>&nbsp;<%=RSLIST("LN")%> From The System."></a></td>
					<td class="m_l_list_t rep_week_t" style="text-align: left;"><%=RSLIST("TITLE")%>. <%=RSLIST("FN")%>&nbsp;<%=RSLIST("LN")%></td>
					<td class="m_l_list_t rep_week_t" style="width: 55px;"><a href="/pt/modules/ss/db/edit.asp?edittype=11&amp;uid=<%=RSLIST("ID")%>&amp;type=2a"><b>Yes</b></a></td>
					<td class="m_l_list_t rep_week_t" style="width: 45px;"><a href="/pt/modules/ss/db/edit.asp?edittype=11&amp;uid=<%=RSLIST("ID")%>&amp;type=2b">Click</a></td>
					<td class="m_l_list_t rep_week_t" style="width: 45px;"><%=RSLIST("ENTITLEMENT")%></td>
					<td class="m_l_list_t rep_week_t" style="width: 45px;"><%=RSLIST("ENTITLEMENT") - RSCOUNT.RECORDCOUNT%></td>
				</tr>
				<%
					elseif (RSLIST(currfield) <> "1") and (RSLIST("D_1") = "1") or (RSLIST("D_2") = "1") or (RSLIST("D_3") = "1") or (RSLIST("D_4") = "1") or (RSLIST("D_5") = "1") or (RSLIST("D_6") = "1") or (RSLIST("D_7") = "1") then
				%>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);">
					<td class="m_l_list_t rep_week_t" style="width: 55px; text-align: center;"><a href="ocover.asp?id=2&amp;uid=<%=RSLIST("ID")%>"><img src="/pt/media/icons/16_edit.gif" border="0" alt="Click Here To Edit <%=RSLIST("FN")%>&nbsp;<%=RSLIST("LN")%>'s Details."></a>&nbsp;<a href="/pt/modules/ss/db/delete.asp?deltype=6&amp;uid=<%=RSLIST("ID")%>"><img src="/pt/media/icons/16_del.gif" border="0" alt="Click Here To Delete <%=RSLIST("FN")%>&nbsp;<%=RSLIST("LN")%> From The System."></a></td>
					<td class="m_l_list_t rep_week_t" style="text-align: left;"><%=RSLIST("TITLE")%>. <%=RSLIST("FN")%>&nbsp;<%=RSLIST("LN")%></td>
					<td class="m_l_list_t rep_week_t" style="width: 55px;"><a href="/pt/modules/ss/db/edit.asp?edittype=11&amp;uid=<%=RSLIST("ID")%>&amp;type=3a">Click</a></td>
					<td class="m_l_list_t rep_week_t" style="width: 55px;"><a href="/pt/modules/ss/db/edit.asp?edittype=11&amp;uid=<%=RSLIST("ID")%>&amp;type=3b"><b>Yes</b></a></td>
					<td class="m_l_list_t rep_week_t" style="width: 45px;"><%=RSLIST("ENTITLEMENT")%></td>
					<td class="m_l_list_t rep_week_t" style="width: 45px;"><%=RSLIST("ENTITLEMENT") - RSCOUNT.RECORDCOUNT%></td>
				</tr>
				<%
					else
				%>
				<tr onMouseOver="colorRowLight(this,1);" onMouseOut="colorRowLight(this,0);">
					<td class="m_l_list_t rep_week_t" style="width: 55px; text-align: center;"><a href="ocover.asp?id=2&amp;uid=<%=RSLIST("ID")%>"><img src="/pt/media/icons/16_edit.gif" border="0" alt="Click Here To Edit <%=RSLIST("FN")%>&nbsp;<%=RSLIST("LN")%>'s Details."></a>&nbsp;<a href="/pt/modules/ss/db/delete.asp?deltype=6&amp;uid=<%=RSLIST("ID")%>"><img src="/pt/media/icons/16_del.gif" border="0" alt="Click Here To Delete <%=RSLIST("FN")%>&nbsp;<%=RSLIST("LN")%> From The System."></a></td>
					<td class="m_l_list_t rep_week_t" style="text-align: left;"><%=RSLIST("TITLE")%>. <%=RSLIST("FN")%>&nbsp;<%=RSLIST("LN")%></td>
					<td class="m_l_list_t rep_week_t" style="width: 55px;"><a href="/pt/modules/ss/db/edit.asp?edittype=11&amp;uid=<%=RSLIST("ID")%>&amp;type=4a">Click</a></td>
					<td class="m_l_list_t rep_week_t" style="width: 55px;"><a href="/pt/modules/ss/db/edit.asp?edittype=11&amp;uid=<%=RSLIST("ID")%>&amp;type=4b">Click</a></td>
					<td class="m_l_list_t rep_week_t" style="width: 45px;"><%=RSLIST("ENTITLEMENT")%></td>
					<td class="m_l_list_t rep_week_t" style="width: 45px;"><%=RSLIST("ENTITLEMENT") - RSCOUNT.RECORDCOUNT%></td>
				</tr>
				<%
				end if
				%>
				<tr>
					<td class="m_l_sel_sep" style="height: 5px;" colspan="5"></td>
				</tr>
				<%
				RSLIST.MOVENEXT
				loop

				end if
				%>
			</table>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td colspan="2" style="height: 10px;"></td>
				</tr>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);" onmouseup="showdetail('add');">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="New Member Of Staff"></td>
					<td class="m_l_list_t"><b>New Outside Cover Member</b></td>
				</tr>
				<tr id="list_add"<%if RSLIST.RECORDCOUNT = 0 then%><%else%> style="display: none;"<%end if%>>
					<td class="m_l_list_p"></td>
					<td class="m_l_list_t" style="padding-top: 7px;">
						<form name="add" action="/pt/modules/ss/db/add.asp?addtype=8" method="post">
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
								<td class="m_l_list_t" style="height: 28px; padding-left: 5px;"><b>Period Entitlement</b></td>
								<td style="font-size: 10pt;">
								<input class="b_std" type="text" name="ENTITLEMENT" size="2" maxlength="2" onblur="noalpha(this.value);"> Periods/Week (Current Week Total Is <b><%=weektotal%></b>)
								</td>
							</tr>
						</table>
						<hr size="1">
						<div style="width: 100%; text-align: center;"><b><a href="#" onmouseup="document.add.submit();">Add</a></b> :: <b><a href="#" onmouseup="document.add.reset();">Clear</a></b></div>
						</form>
					</td>
				</tr>
			</table>
			<hr size="1">
			<div style="width: 100%; text-align: right; font-size: 18pt; font-weight: bold;"><a href="cover.asp?id=1">Arrange Staff Cover></a></div>
		</div>
	</div>
<%
	RSLIST.CLOSE
	set RSLIST = nothing
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