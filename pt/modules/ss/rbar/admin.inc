<h2 class="nocss">Further Information/Links</h2>
<%
RSCLASSESCOVEREDSQL = "SELECT * FROM COVER WHERE COVERING = " & session("sess_ttid")

Set RSCLASSESCOVERED = Server.CreateObject("Adodb.RecordSet")
RSCLASSESCOVERED.Open RSCLASSESCOVEREDSQL, dataconn, adopenkeyset, adlockoptimistic

if session("sess_adminlevel") <> "1" then
%>
<div class="m_r_ti">Your Information</div>
<div class="m_r_t">
<b><%=session("sess_un")%></b>&nbsp;(<%=session("sess_fn")%>&nbsp;<%=session("sess_ln")%>)<br>
<b>Limited Access</b> Account
</div>
<div class="m_r_ti">QuickLinks</div>
<div class="m_r_t">
	<a href="cover.asp?id=1"><b>Arrange Cover For Today</b></a><br>
	<a href="staff.asp?id=3">View Your Staff's Timetables</a><br>
	<a href="reports.asp?id=3">View Today's Cover Summary</a><br>
</div>
<%
else
%>
<div class="m_r_ti">Today's Cover Summary</div>
<div class="m_r_t">
<!--#include virtual="/pt/modules/ss/timetables/admin_rbar_sum.inc"-->
</div>

<%
if session("sess_un") = "MaxwellD1" then
%>
<div class="m_r_ti">Remaining Tasks</div>
<div class="m_r_t">
	<b>Complete Settings</b><br>
	<b>Complete Staff Mgt (Properly)</b><br>
	<b>Complete Room Mgt</b><br>
	<b>Complete Dept. Mgt</b><br>
	<b>JOB SHARES</b><br>
	<b>Complete Backup</b>
		<div style="padding-right: 10px;">
		Backup Staff Details<br>
		Weekly Report For B-U-D
		</div>
	<b>Complete Reports</b>
		<div style="padding-right: 10px;">
		Staff Details
		</div>	
	<b>Signup?</b><br>
	<b>Update Setup Script</b><br>
	<b>Fix PIN Bug (Login)</b><br>
	<b>User Acc. Update Fix</b><br>
	<b>When Delete User</b>
		<div style="padding-right: 10px;">
		Delete Cover<br>
		Del. Attendance<br>
		?Backup Info?
		</div>	
</div>
<%
else
end if
%>

<div class="m_r_ti">QuickLinks</div>
<div class="m_r_t">
	<a href="cover.asp?id=1"><b>Arrange Staff Cover</b></a><br>
	<a href="ocover.asp?id=1">Arrange Outside Cover</a><br>
</div>

<div class="m_r_t">
	<a href="staff.asp?id=6">Add A Member Of Staff</a><br>
	<a href="staff.asp?id=5">Edit Someone's Details</a><br>
	<a href="staff.asp?id=3">Edit Someone's Timetable</a><br>
</div>

<div class="m_r_t">
	<a href="reports.asp?id=2&amp;print=1">Print Today's PleaseTake Slips</a><br>
	<a href="reports.asp?id=3">View Today's Cover Summary</a><br>
</div>

<div class="m_r_t">
	<a href="settings.asp?id=1">Change System Settings</a><br>
	<a href="backup.asp?id=2">Backup Data</a><br>
</div>

<div class="m_r_ti">About The System</div>
<div class="m_r_t">
	<img alt="PleaseTakes V1" src="/pt/media/admin/pt_about.png">
	<b>Version <%=var_ver%></b><br>
	For <%=var_est_full%><br>
	<b>&copy;&nbsp;<a href="http://www.server-ml.co.uk/" target="_blank">Server-ML.co.uk</a> 2006</b>
</div>
<%
end if

RSCLASSESCOVERED.CLOSE
set RSCLASSESCOVERED = nothing
%>