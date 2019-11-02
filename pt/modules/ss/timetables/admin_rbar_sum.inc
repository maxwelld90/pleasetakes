<%
'Find Out Maximum Number Of Periods

RSMAXSQL = "SELECT MAX(Totals) AS Highest FROM Periods"

Set RSMAX = Server.CreateObject("Adodb.RecordSet")
RSMAX.Open RSMAXSQL, dataconn, adopenkeyset, adlockoptimistic

'Retrieve Individual Days, Put Into Array

RSDAYSQL = "SELECT TOTALS FROM Periods WHERE DAYID = " & DOW_SUM

Set RSDAY = Server.CreateObject("Adodb.RecordSet")
RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic

daydate = date()
%>
<table class="m_l_tt" style="height: 60px;" cellspacing="0" cellpadding="0">
	<tr>
		<%
		for i=1 to RSMAX("Highest")
		%>
		<td class="m_l_tt_td" style="height: 25px; font-family: Tahoma,Sans-Serif; font-size: 10pt;"><%=i%></td>
		<%
		next
		%>
	</tr>
	<tr>
		<%
		daytot = RSMAX("Highest") - RSDAY("TOTALS")

		if daytot = RSMAX("Highest") then
		%>
		<td class="m_l_tt_td" style="_width: 100%; height: 25px; font-family: Tahoma,Sans-Serif; font-size: 10pt; background-image: url('/pt/media/tt/r.png'); background-repeat: repeat-x;" colspan="<%=RSMAX("Highest")%>">No Periods Today!</td>
		<%
		else

			for i=1 to RSDAY("TOTALS")

			RSCOVERCOUNTSQL = "SELECT Cover.ID, Cover.FOR, Cover.COVERING, Cover.DAY, Cover.DAYDATE, Cover.PERIOD, Cover.OCOVER FROM Cover WHERE (((Cover.DAY)=" & DOW_SUM & ") AND ((Cover.DAYDATE)=#" & SQLDATE(date()) & "#) AND ((Cover.PERIOD)=" & i & "));"

			Set RSCOVERCOUNT = Server.CreateObject("Adodb.RecordSet")
			RSCOVERCOUNT.Open RSCOVERCOUNTSQL, dataconn, adopenkeyset, adlockoptimistic
	
		%>
		<td class="m_l_tt_td" style="height: 25px; font-family: Tahoma,Sans-Serif; font-size: 8pt; font-weight: normal; background-image: url('/pt/media/tt/g.png'); background-repeat: repeat-x;"><%=RSCOVERCOUNT.RECORDCOUNT%></td>
		<%
			RSCOVERCOUNT.CLOSE
			set RSCOVERCOUNT = nothing

			next

		for i=1 to daytot
		%>
		<td class="m_l_tt_td" style="height: 25px; font-family: Tahoma,Sans-Serif; font-size: 10pt; background-image: url('/pt/media/tt/r.png'); background-repeat: repeat-x;"></td>
		<%
		next
		end if
		set daytot = nothing
		%>		
	</tr>
</table>