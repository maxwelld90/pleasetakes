	<h1 class="nocss">Server-ML.co.uk <%=var_pname%> Version <%=var_ver%></h1>
	<h2 class="nocss">Your Details/Account Options</h2>
	<div class="smlb_k"></div>
	<div class="smlb_l">You Are Logged In As <b><%=session("sess_fn")%>&nbsp;<%=session("sess_ln")%></b></div>
	<div class="smlb_r">
		<ul>
			<li><a href="settings.asp?id=2" title="Click Here To Change Your Account Settings.">Change My Settings</a> |</li>		
			<li><a href="/pt/modules/ss/usersys/logout.asp?id=2" title="Click Here To Logout Of The System.">Logout</a></li>
		</ul>
	</div>
	<div class="topb_t"><%=var_pname%> Admin</div>
	<h2 class="nocss">Main Menu</h2>
	<div class="topb_m">
		<ul>
			<%
			if session("sess_adminlevel") <> 1 then
			%>
			<li><a href="default.asp">Home</a> :: </li>
			<%
				if session("sess_dept") <> "" then
			%>
			<li><a href="cover.asp?id=1"><b>Arrange Departmental Cover</b></a> :: </li>
			<li><a href="staff.asp?id=3">Your Staff Timetables</a> :: </li>
			<%
				else
				end if
			%>
			<li><a href="reports.asp?id=3">Today's Cover Summary</a> :: </li>
			<li><a href="/pt/modules/ss/usersys/logout.asp?id=2" title="Click Here To Logout Of The System.">Logout</a></li>
			<%
			else
			%>
			<li><a href="default.asp">Home</a> :: </li>
			<li><a href="cover.asp?id=1"><b>Arrange Staff Cover</b></a> :: </li>
			<li><a href="staff.asp?id=1">Staff Mgt.</a> :: </li>
			<li><a href="rooms.asp?id=1">Room Mgt.</a> :: </li>
			<li><a href="ocover.asp?id=1">Outside Cover</a> :: </li>
			<li><a href="reports.asp?id=1">Reports</a> :: </li>
			<li><a href="backup.asp?id=1">Data Backup</a> :: </li>
			<li><a href="settings.asp?id=1">System Management</a> :: </li>
			<li><a href="/pt/modules/ss/usersys/logout.asp?id=2" title="Click Here To Logout Of The System.">Logout</a></li>
			<%
			end if
			%>
		</ul>
	</div>
	<h2 class="nocss">Date And Time</h2>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>
	<h2 class="nocss">Page Content</h2>