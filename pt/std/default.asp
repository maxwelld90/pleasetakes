<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" >
<!--#include virtual="/pt/modules/ss/usersys/logincheck_std.inc"-->

<!--#include virtual="/pt/modules/ss/p_s.inc"-->

<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="../modules/css/std.css">
<script language="javascript" type="text/javascript" src="/pt/modules/js/std.js"></script>
<title><%=var_ptitle%></title>
</head>

<body>

<div class="smlb_b"></div>
<div class="topb_b"></div>

<div class="main">
	<!--#include virtual="/pt/modules/ss/topbar/std.inc"-->

	<div class="m_l">
		<div class="m_l_title"><script language="javascript" type="text/javascript">document.write(daymsg)</script>, <%=session("sess_fn")%>!</div>
		<div class="m_l_subtitle">Welcome To The System!</div>
		<div class="m_l_ins">Please Choose A Task From Below...</div>
		<div class="m_l_sel">
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_sel_t_l"><a href="view.asp?id=2"><img src="/pt/media/icons/48_view.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="view.asp?id=2">View My PleaseTakes</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="view.asp?id=3"><img src="/pt/media/icons/48_cal.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="view.asp?id=3">View My Timetable</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="#"><img src="/pt/media/icons/48_setting.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="#">Change My Settings</a></td>
				</tr>
				<tr>
					<td class="m_l_sel_sep" colspan="2"></td>
				</tr>
				<tr>
					<td class="m_l_sel_t_l"><a href="/pt/modules/ss/usersys/logout.asp?id=1"><img src="/pt/media/icons/48_key.png" border="0"></a></td>
					<td class="m_l_sel_t_r"><a href="/pt/modules/ss/usersys/logout.asp?id=1">Logout</a></td>
				</tr>
			</table>
		</div>
	</div>
	<div class="m_r">
		<div class="m_r2">
		<!--#include virtual="/pt/modules/ss/rbar/std.inc"-->
		</div>
	</div>
</div>

</body>

</html>

<!--#include virtual="/pt/modules/ss/p_e.inc"-->