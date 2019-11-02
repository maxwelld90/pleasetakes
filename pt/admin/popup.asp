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

<div class="main" style="width: 634px;">
	<!--#include virtual="/pt/modules/ss/topbar/admin_popup.inc"-->
<%
if pagetype = "1" then

	if (request("err")) = "1" then
%>
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
	<form name="edit" action="/pt/modules/ss/db/edit.asp?edittype=2&amp;period=<%=request("period")%>&amp;day=<%=request("day")%>" method="post">
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Staff Timetable</div>
		<div class="m_l_subtitle">Editing A Period</div>
		<div class="m_l_ins">Please Change The Settings Below For This Period, Then Click "Save".</div>
		<div class="m_l_sel">
		<%
		RSUSERSQL = "SELECT * FROM Timetables WHERE ID = " & request("user")
		RSROOMSQL = "SELECT * FROM Rooms ORDER BY ROOMNO"
		
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
											if RSUSER("DEFROOM") <> "" then
											%>
												<option value="DEF">Usual Room</option>
											<%
											else
											end if
											%>
												<option value="NA"<%if RSUSER("R" & request("period") & "_" & request("day")) = "NA" then%> selected<%else%><%end if%>>N/A (No Room)</option>
											<%
											do until RSROOM.EOF
											%>
												<option<%if RSROOM("ROOMNO") = RSUSER("R" & request("period") & "_" & request("day")) then%> selected<%else%><%end if%>><%=RSROOM("ROOMNO")%></option>
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

elseif pagetype = "2" then

	RSCHECKSQL = "SELECT * FROM Inventory WHERE WeekNo = " & request("wkno") & " AND BackupYear = " & request("year")
	
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, backupconn, adopenkeyset, adlockoptimistic

	if RSCHECK.RECORDCOUNT = 0 then
		RSCHECK.close
		set RSCHECK = nothing
		response.write "<script language='javascript' type='text/javascript'>opener.location.reload(true); self.close();</script>"
	else

		RSWEEKINFOSQL = "SELECT * FROM [C_" & RSCHECK("StartDate") & "_" & RSCHECK("EndDate") & "_" & RSCHECK("WeekNo") & "] WHERE Day = " & request("DOW") & " AND Period = " & request("period")
	
		Set RSWEEKINFO = Server.CreateObject("Adodb.RecordSet")
		RSWEEKINFO.Open RSWEEKINFOSQL, backupconn, adopenkeyset, adlockoptimistic

		if (request("DOW")) = "1" then
			backupdaydate = RSHCHECK("StartDate")
		elseif (request("DOW")) = "2" then
			backupdaydate = (DateAdd("d",1,RSCHECK("StartDate")))
		elseif (request("DOW")) = "3" then
			backupdaydate = (DateAdd("d",2,RSCHECK("StartDate")))
		elseif (request("DOW")) = "4" then
			backupdaydate = (DateAdd("d",3,RSCHECK("StartDate")))
		elseif (request("DOW")) = "5" then
			backupdaydate = (DateAdd("d",4,RSCHECK("StartDate")))
		elseif (request("DOW")) = "6" then
			backupdaydate = (DateAdd("d",5,RSCHECK("StartDate")))
		elseif (request("DOW")) = "7" then
			backupdaydate = (DateAdd("d",6,RSCHECK("StartDate")))
		end if

	if RSWEEKINFO.RECORDCOUNT = 0 then
	%>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">No Staff Absent!</div>
		<div class="m_l_subtitle">Everyone Was Present On <%=GetDOW(request("DOW"))%>, Period <%=request("Period")%>!</div>
		<div class="m_l_ins">Must Have Been A Good Day!</div>
		<div class="m_l_sel">
		No Absent Staff Were Found For This Period And Day!<br>
		Please Click <a href="#" onmouseup="self.close();">Here</a> To Close This Popup Window.
		</div>
	</div>
<%
	else
%>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Who Was Off?</div>
		<div class="m_l_subtitle">See Who Was Off On <%=GetDOW(request("DOW"))%>, <%=backupdaydate%>, Period <%=request("Period")%>.</div>
		<div class="m_l_ins">The Absent Staff Members And The Covering Staff Are Listed Below.</div>
		<div class="m_l_sel">

		No Absent Staff Were Found For This Period And Day!<br>
		Please Click <a href="#" onmouseup="self.close();">Here</a> To Close This Popup Window.
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_sep" colspan="3"></td>
				</tr>
				<%
				do until RSWEEKINFO.EOF

					RSCHECK2SQL = "SELECT * FROM Timetables WHERE ID = " & RSWEEKINFO("For")
	
					Set RSCHECK2 = Server.CreateObject("Adodb.RecordSet")
					RSCHECK2.Open RSCHECK2SQL, dataconn, adopenkeyset, adlockoptimistic
					
						if RSCHECK2.RECORDCOUNT = 0 then
							fn_for = "A Staff Member Who Has Left"
						else
							fn_for = RSCHECK2("FN") & "&nbsp;"
							ln_for = RSCHECK2("LN")
						end if
					
					RSCHECK2.close
					set RSCHECK2 = nothing

					RSCHECK3SQL = "SELECT * FROM Timetables WHERE ID = " & RSWEEKINFO("Covering")
	
					Set RSCHECK3 = Server.CreateObject("Adodb.RecordSet")
					RSCHECK3.Open RSCHECK3SQL, dataconn, adopenkeyset, adlockoptimistic

						if RSCHECK3.RECORDCOUNT = 0 then
							fn_covering = "A Staff Member Who Has Left"
						else
							fn_covering = RSCHECK3("FN") & "&nbsp;"
							ln_covering = RSCHECK3("LN")
						end if
					
					RSCHECK3.close
					set RSCHECK3 = nothing
				%>
				<tr class="m_l_list_b" onMouseOver="colorRow(this,1);" onMouseOut="colorRow(this,0);">
					<td class="m_l_list_p"><img src="/pt/media/icons/16_man.gif" border="0" alt="<%=fn_for%><%=ln_for%> Was Absent."></td>
					<td class="m_l_list_t"><b><%=fn_for%><%=ln_for%></b> Was Absent, And Was Covered By <b><%=fn_covering%><%=ln_covering%></b>.</td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="3"></td>
				</tr>
				<%
				RSWEEKINFO.MOVENEXT
				loop
				%>
			</table>
			<hr size="1">
			<div style="width: 100%; text-align: right;">Displaying A Total Of <b><%=RSWEEKINFO.RECORDCOUNT%></b> Absent Staff Member<%if RSWEEKINFO.RECORDCOUNT = 1 then%><%else%>s<%end if%></div>
		</div>
	</div>
<%
	end if

	RSCHECK.close
	set RSCHECK = nothing

	end if

elseif pagetype = "3" then
	if (request("confirm") = "1") then
%>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Are You Sure?</div>
		<div class="m_l_subtitle">You Will Have To Manually Reselect Cover For This Period.</div>
		<div class="m_l_ins">Is It A Yes Or Is It A No?</div>
		<div class="m_l_sel">
			Please Confirm If You Really Wish To Delete This Cover Request.
			Clicking "Yes!" Deletes And Closes The Popup Window, And "No!" Takes You Back.<br>
			<div style="font-size: 56pt; letter-spacing: -5px; font-weight: bold; text-align: center;"><a href="#" onmouseup="location.href='/pt/modules/ss/db/delete.asp?deltype=8&amp;cover=<%=request("cover")%>'">Yes!</a></div>
			<div style="font-size: 56pt; letter-spacing: -5px; font-weight: bold; text-align: center;"><a href="#" onmouseup="history.back();">No!</a></div>
		</div>
	</div>
<%
	else

	daydow = request("daydow")
	daydate = request("daydate")

			RSCHECKSQL = "SELECT * FROM Cover WHERE ID = " & request("cover")
		
			Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
			RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
			
			if RSCHECK("OCOVER") = "1" then
				RSCOVERSQL = "SELECT * FROM Cover WHERE FOR = " & RSCHECK("FOR") & " AND DAY = " & daydow & " AND DAYDATE = #" & SQLDate(daydate) & "#" 
				Set RSCOVER = Server.CreateObject("Adodb.RecordSet")
				RSCOVER.Open RSCOVERSQL, dataconn, adopenkeyset, adlockoptimistic

				RSCLASSSQL = "SELECT [" & RSCHECK("PERIOD") & "_" & daydow & "], [R" & RSCHECK("PERIOD") & "_" & daydow & "] FROM Timetables WHERE ID = " & RSCHECK("FOR")	
				Set RSCLASS = Server.CreateObject("Adodb.RecordSet")
				RSCLASS.Open RSCLASSSQL, dataconn, adopenkeyset, adlockoptimistic

				RSFORSQL = "SELECT * FROM Timetables WHERE ID = " & RSCHECK("FOR")	
				Set RSFOR = Server.CreateObject("Adodb.RecordSet")
				RSFOR.Open RSFORSQL, dataconn, adopenkeyset, adlockoptimistic

					RSFORDEPTSQL = "SELECT * FROM Departments WHERE DEPTID = " & RSFOR("DEPT")		
					Set RSFORDEPT = Server.CreateObject("Adodb.RecordSet")
					RSFORDEPT.Open RSFORDEPTSQL, dataconn, adopenkeyset, adlockoptimistic		
		
				RSCOVERINGSQL = "SELECT * FROM OCover WHERE ID = " & RSCHECK("COVERING")
				Set RSCOVERING = Server.CreateObject("Adodb.RecordSet")
				RSCOVERING.Open RSCOVERINGSQL, dataconn, adopenkeyset, adlockoptimistic
			else
				RSCOVERSQL = "SELECT * FROM Cover WHERE FOR = " & RSCHECK("FOR") & " AND DAY = " & daydow & " AND DAYDATE = #" & SQLDate(daydate) & "#" 
				Set RSCOVER = Server.CreateObject("Adodb.RecordSet")
				RSCOVER.Open RSCOVERSQL, dataconn, adopenkeyset, adlockoptimistic

				RSCLASSSQL = "SELECT [" & RSCHECK("PERIOD") & "_" & daydow & "], [R" & RSCHECK("PERIOD") & "_" & daydow & "] FROM Timetables WHERE ID = " & RSCHECK("FOR")	
				Set RSCLASS = Server.CreateObject("Adodb.RecordSet")
				RSCLASS.Open RSCLASSSQL, dataconn, adopenkeyset, adlockoptimistic

				RSFORSQL = "SELECT * FROM Timetables WHERE ID = " & RSCHECK("FOR")	
				Set RSFOR = Server.CreateObject("Adodb.RecordSet")
				RSFOR.Open RSFORSQL, dataconn, adopenkeyset, adlockoptimistic

					RSFORDEPTSQL = "SELECT * FROM Departments WHERE DEPTID = " & RSFOR("DEPT")		
					Set RSFORDEPT = Server.CreateObject("Adodb.RecordSet")
					RSFORDEPT.Open RSFORDEPTSQL, dataconn, adopenkeyset, adlockoptimistic		
		
				RSCOVERINGSQL = "SELECT * FROM Timetables WHERE ID = " & RSCHECK("COVERING")
				Set RSCOVERING = Server.CreateObject("Adodb.RecordSet")
				RSCOVERING.Open RSCOVERINGSQL, dataconn, adopenkeyset, adlockoptimistic
	
					RSCOVERINGDEPTSQL = "SELECT * FROM Departments WHERE DEPTID = " & RSCOVERING("DEPT")
					Set RSCOVERINGDEPT = Server.CreateObject("Adodb.RecordSet")
					RSCOVERINGDEPT.Open RSCOVERINGDEPTSQL, dataconn, adopenkeyset, adlockoptimistic	
			end if
%>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Cover Info & Options</div>
		<div class="m_l_subtitle">View More Information On This Request, Or Perform Tasks.</div>
		<div class="m_l_ins">Please Select An Option From Below.</div>
		<div class="m_l_sel">
			<div class="m_l_middletitle" style="width: 630px; margin-bottom: 15px; font-size: 12pt;" onmouseup="showdetail('info');">View Detailed Information For This Cover Request</div>
			<div id="list_info" style="padding: 0px 0px 5px 10px; margin-top: -10px; display: none;">
			<span style="font-size: 12pt; font-weight: bold; letter-spacing: -1px;">Cover Type...</span> <%if RSCHECK("OCover") = "1" then%>Outside Cover<%else%>Teaching Staff Cover<%end if%><br>
			<span style="font-size: 12pt; font-weight: bold; letter-spacing: -1px;">Who's Off...</span> <%=RSFOR("FN")%>&nbsp;<%=RSFOR("LN")%> (<%=RSFORDEPT("Full")%>)<br>
			<span style="font-size: 12pt; font-weight: bold; letter-spacing: -1px;">Who's Covering...</span> <%=RSCOVERING("FN")%>&nbsp;<%=RSCOVERING("LN")%><%if RSCHECK("OCover") = "1" then%><%else%>&nbsp;(<%=RSCOVERINGDEPT("FULL")%>)<%end if%><br>
			<span style="font-size: 12pt; font-weight: bold; letter-spacing: -1px;">When?</span> <%=getdow(daydow)%>, <%=day(daydate)%>&nbsp;<%=cal_getmonth(month(daydate))%>&nbsp;<%=year(daydate)%>, Period <%=RSCHECK("Period")%><br>
			<span style="font-size: 12pt; font-weight: bold; letter-spacing: -1px;">Request ID...</span> <%=RSCHECK("ID")%><br>
			</div>
			<div class="m_l_middletitle" style="width: 630px; margin-bottom: 15px; font-size: 12pt;" onmouseup="showdetail('add');">Perform Additional Tasks</div>
			<div id="list_add" style="padding: 0px 0px 5px 10px; margin-top: -10px; font-size: 12pt; font-weight: bold;">
			<a href="#" onmouseup="location.href='print.asp?id=8&amp;cover=<%=request("cover")%>&amp;daydow=<%=request("daydow")%>&amp;daydate=<%=request("daydate")%>'">Print The Individual PleaseTake Slip For This Cover Request</a>
				<div style="padding-left: 20px; font-family: Tahoma, Sans-Serif; font-size: 8pt; font-weight: normal;">
				Click Above To Send The PleaseTake Slip For This Request To Your Printer, Which Saves You Lots And Lots Of Paper.
				</div>
			<a href="#" onmouseup="location.href='popup.asp?id=3&amp;cover=<%=request("cover")%>&amp;confirm=1'">Delete This Cover Request</a>
				<div style="padding-left: 20px; font-family: Tahoma, Sans-Serif; font-size: 8pt; font-weight: normal;">
				Deletes The Cover Request, Leaving The Selected Period Uncovered.
				</div>
			<a href="#" onmouseup="self.close();">Close This Window</a>
				<div style="padding-left: 20px; font-family: Tahoma, Sans-Serif; font-size: 8pt; font-weight: normal;">
				Closes This Popup Window And Returns You To The Cover Summary Page.
				</div>
			</div>
		</div>
	</div>
<%
		if RSCHECK("OCover") = "1" then
		else
		RSCOVERINGDEPT.close
		end if
		RSCOVERING.close
		RSFORDEPT.close
		RSFOR.close
		RSCLASS.close
		RSCOVER.close
		if RSCHECK("OCover") = "1" then
		else
		set RSCOVERINGDEPT = nothing
		end if
		set RSCOVERING = nothing
		set RSFORDEPT = nothing
		set RSFOR = nothing
		set RSCLASS = nothing
		set RSCOVER = nothing
		RSCHECK.close
		set RSCHECK = nothing
	end if
else
%>
	<div class="m_l" style="width: 100%;">
		<div class="m_l_title">Popup Error</div>
		<div class="m_l_subtitle">No Page ID Specified</div>
		<div class="m_l_ins">The Popup Cannot Be Loaded.</div>
		<div class="m_l_sel">
			Sorry, <%=session("sess_fn")%>, But This Popup Cannot Be Loaded Because No Page ID Has Been Specified.
			Please Go Back And Try Another Option.
			<div class="botopts">
				<ul>
					<li><a href="#" onmouseup="history.back();">Return</a></li>
					<li><a href="#" onmouseup="self.close();">Close The Popup Window</a></li>
				</ul>
			</div>
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