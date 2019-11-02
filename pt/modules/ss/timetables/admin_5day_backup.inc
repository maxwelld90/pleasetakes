<%
'Find Out Maximum Number Of Periods

RSMAXSQL = "SELECT MAX(Totals) AS Highest FROM Periods"

Set RSMAX = Server.CreateObject("Adodb.RecordSet")
RSMAX.Open RSMAXSQL, dataconn, adopenkeyset, adlockoptimistic

'Retrieve Individual Days, Put Into Array

RSDAYSQL = "SELECT * FROM Periods"

Set RSDAY = Server.CreateObject("Adodb.RecordSet")
RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic

Dim dayarray2(6)

for i=1 to 6
dayarray2(i) = RSDAY("TOTALS")
RSDAY.MOVENEXT
next
RSDAY.MOVEFIRST
%>
<table class="m_l_tt" cellspacing="0" cellpadding="0">
	<tr>
		<td class="m_l_tt_td"></td>
		<%
		for i=1 to RSMAX("Highest")
		%>
		<td class="m_l_tt_td" style="font-size: 13pt;"><%=i%></td>
		<%
		next
		%>
	</tr>
	<%
	RSDAY.MOVENEXT
	for j = 2 to 6
	%>
	<tr>
		<td class="m_l_tt_td" style="font-size: 13pt;">
		<%
		if j = 2 then
		response.write RSDAY("Dayname")
		RSDAY.MOVENEXT
		elseif j = 3 then
		response.write RSDAY("Dayname")
		RSDAY.MOVENEXT
		elseif j = 4 then
		response.write RSDAY("Dayname")
		RSDAY.MOVENEXT
		elseif j = 5 then
		response.write RSDAY("Dayname")
		RSDAY.MOVENEXT
		elseif j = 6 then
		response.write RSDAY("Dayname")
		RSDAY.MOVENEXT
		end if
		%>
		</td>
		<%
		daytot = RSMAX("Highest") - dayarray2(j)

		if daytot = RSMAX("Highest") then
			for i=1 to RSMAX("Highest")
		%>
		<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/r.png'); background-repeat: repeat-x;"></td>
		<%
		next
		
		else

			for i=1 to dayarray2(j)

			RSWEEKINFOSQL = "SELECT * FROM [C_" & RSWEEK("StartDate") & "_" & RSWEEK("EndDate") & "_" & RSWEEK("WeekNo") & "] WHERE Period = " & i & " AND Day = " & j
	
			Set RSWEEKINFO = Server.CreateObject("Adodb.RecordSet")
			RSWEEKINFO.Open RSWEEKINFOSQL, backupconn, adopenkeyset, adlockoptimistic
		%>
		<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/g.png'); background-repeat: repeat-x;" onmouseup="popup('popup.asp?id=2&amp;period=<%=i%>&amp;dow=<%=j%>&amp;wkno=<%=request("week")%>&amp;year=<%=request("year")%>');"><%=RSWEEKINFO.RECORDCOUNT%></td>
		<%
			RSWEEKINFO.close
			set RSWEEKINFO = nothing
			
			next
		k = 0
		per = i
		do until k=daytot
		%>
		<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/r.png'); background-repeat: repeat-x;"></td>
		<%
		per = per + 1
		k = k + 1
		loop

		end if
		set daytot = nothing
		%>		
	</tr>
	<%
	next
	%>
</table>