<%
if var_est_enabled <> "1" then
%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t"><b>The System Is Currently Disabled! Do You Want To <a href="settings.asp?id=3">Enable It</a>?</b></td>
				</tr>
			</table>
			<hr size="1">
<%
else
end if
%>