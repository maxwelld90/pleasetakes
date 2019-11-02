<%
'Find Out Maximum Number Of Periods

RSMAXSQL = "SELECT MAX(Totals) AS Highest FROM Periods"

Set RSMAX = Server.CreateObject("Adodb.RecordSet")
RSMAX.Open RSMAXSQL, dataconn, adopenkeyset, adlockoptimistic

'Retrieve Individual Days, Put Into Array

RSDAYSQL = "SELECT TOTALS FROM Periods WHERE DAYID = " & DOW

Set RSDAY = Server.CreateObject("Adodb.RecordSet")
RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic
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

			RSCOVERCOUNTSQL = "SELECT * FROM Cover WHERE PERIOD = " & i & " AND DAY = " & DOW & " AND COVERING = " & session("sess_ttid")

			Set RSCOVERCOUNT = Server.CreateObject("Adodb.RecordSet")
			RSCOVERCOUNT.Open RSCOVERCOUNTSQL, dataconn, adopenkeyset, adlockoptimistic
		
		if RSCOVERCOUNT.RECORDCOUNT => 1 then	
		%>
		<td class="m_l_tt_td" style="height: 25px; font-family: Tahoma,Sans-Serif; font-size: 8pt; background-image: url('/pt/media/tt/g.png'); background-repeat: repeat-x;"></td>
		<%
		else
		%>
	<td class="m_l_tt_td" style="height: 25px; font-family: Tahoma,Sans-Serif; font-size: 8pt; background-image: url('/pt/media/tt/na.png'); background-repeat: repeat-x;"></td>
		<%
		end if
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
Green Means You Have To Cover A Class<br>
Grey Means You Don't