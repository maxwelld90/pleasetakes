<%
'Find Out Maximum Number Of Periods

RSMAXSQL = "SELECT MAX(Totals) AS Highest FROM Periods"

Set RSMAX = Server.CreateObject("Adodb.RecordSet")
RSMAX.Open RSMAXSQL, dataconn, adopenkeyset, adlockoptimistic

'Retrieve Individual Days, Put Into Array

RSDAYSQL = "SELECT * FROM Periods"

Set RSDAY = Server.CreateObject("Adodb.RecordSet")
RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic

Dim dayarray(7)

for i=1 to 7
dayarray(i) = RSDAY("TOTALS")
RSDAY.MOVENEXT
next
RSDAY.MOVEFIRST
%>
<table class="m_l_tt" cellspacing="0" cellpadding="0">
	<tr>
		<td class="m_l_tt_td"></td>
		<%
		for i=1 to var_maxperiods
		%>
		<td class="m_l_tt_td" style="font-size: 13pt;" onmouseup="location.href='/pt/modules/ss/db/setup.asp?setuptype=5&amp;period=<%=i%>';"><%=i%></td>
		<%
		next
		%>
	</tr>

	<%
	for j = 1 to 7
	%>
	<tr>
		<td class="m_l_tt_td" style="font-size: 13pt;" onmouseup="location.href='/pt/modules/ss/db/setup.asp?setuptype=5&amp;dow=<%=j%>';">
		<%
		if j = 1 then
		response.write RSDAY("Dayname")
		RSDAY.MOVENEXT
		elseif j = 2 then
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
		elseif j = 7 then
		response.write RSDAY("Dayname")
		RSDAY.MOVENEXT
		end if
		%>
		</td>
		<%
		daytot = var_maxperiods - dayarray(j)

		if daytot = var_maxperiods then
			for i=1 to var_maxperiods
		%>
		<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/r.png'); background-repeat: repeat-x;" onmouseup="location.href='/pt/modules/ss/db/setup.asp?setuptype=5&amp;period=<%=i%>&amp;dow=<%=j%>';"></td>
		<%
		next
		
		else

			for i=1 to dayarray(j)
		%>
		<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/g.png'); background-repeat: repeat-x;" onmouseup="location.href='/pt/modules/ss/db/setup.asp?setuptype=5&amp;period=<%=i%>&amp;dow=<%=j%>';"></td>
		<%
			next
		k = 0
		per = i
		do until k=daytot
		%>
		<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/r.png'); background-repeat: repeat-x;" onmouseup="location.href='/pt/modules/ss/db/setup.asp?setuptype=5&amp;period=<%=per%>&amp;dow=<%=j%>';"></td>
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