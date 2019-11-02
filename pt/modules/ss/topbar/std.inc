	<div class="smlb_k"></div>
	<div class="smlb_l">You Are Logged In As <b><%=session("sess_fn")%>&nbsp;<%=session("sess_ln")%></b></div>
	<div class="smlb_r">
		<ul>
			<li><a href="#">Change My Settings</a> |</li>		
			<li><a href="/pt/modules/ss/usersys/logout.asp?id=1" title="Click Here To Logout Of The System.">Logout</a></li>
		</ul>
	</div>
	<div class="topb_t"><%=var_pname%></div>
	<div class="topb_m">
		<ul>
			<li><a href="default.asp?id=1">Home</a> :: </li>
			<li><a href="view.asp?id=2">View My PleaseTakes</a> :: </li>
			<li><a href="view.asp?id=3">View My Timetable</a> :: </li>
			<li><a href="#">Change My Settings</a> :: </li>
			<li><a href="/pt/modules/ss/usersys/logout.asp?id=1" title="Click Here To Logout Of The System.">Logout</a></li>
		</ul>
	</div>
	<div class="topb_dt"><%=display_date%>, <%=display_time%></div>