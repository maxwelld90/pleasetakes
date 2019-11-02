<%
RSCHECKSQL = "SELECT CATEGORY FROM Timetables"
Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic

count = RSCHECK.RECORDCOUNT

if count = 0 then
%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Have No Staff! Click <a href="staff.asp?id=6"><b>Here</b></a> To Add Some!</td>
				</tr>
			</table>
			<hr size="1">
<%
RSCHECK.close
set RSCHECK = nothing

else
	RSCHECKSQL = "SELECT CATEGORY FROM Timetables WHERE CATEGORY = 'HT'"
	Set RSCHECK = Server.CreateObject("Adodb.RecordSet")
	RSCHECK.Open RSCHECKSQL, dataconn, adopenkeyset, adlockoptimistic
	if RSCHECK.RECORDCOUNT = 0 then
%>
			<table class="m_l_sel_t" cellpadding="0" cellspacing="0">
				<tr>
					<td class="m_l_list_p"><img src="/pt/media/icons/32_alert.png" border="0" alt="Alert!"></td>
					<td class="m_l_list_t">You Have No Head Teacher! Click <a href="staff.asp?id=6"><b>Here</b></a> To Add One!</td>
				</tr>
			</table>
			<hr size="1">
<%
else
end if

end if

%>