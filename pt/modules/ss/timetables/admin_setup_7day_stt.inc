<%
RSUSERSQL = "SELECT * FROM Timetables WHERE ID = " & RSUSERSETUP("ID")
	
Set RSUSER = Server.CreateObject("Adodb.RecordSet")
RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic

'Find Out Maximum Number Of Periods

RSMAXSQL = "SELECT MAX(Totals) AS Highest FROM Periods"

Set RSMAX = Server.CreateObject("Adodb.RecordSet")
RSMAX.Open RSMAXSQL, dataconn, adopenkeyset, adlockoptimistic

'Retrieve Individual Days, Put Into Array

RSDAYSQL = "SELECT * FROM Periods"

Set RSDAY = Server.CreateObject("Adodb.RecordSet")
RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic

Redim dayarray(7)

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
		for i=1 to RSMAX("Highest")
		%>
		<td class="m_l_tt_td" style="font-size: 13pt;"><%=i%></td>
		<%
		next
		%>
	</tr>

	<%
	for j = 1 to 7
	%>
	<tr>
		<td class="m_l_tt_td" style="font-size: 13pt;">
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
		daytot = RSMAX("Highest") - dayarray(j)

		if daytot = RSMAX("Highest") then
			for i=1 to RSMAX("Highest")
		%>
		<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/na_a.png'); background-repeat: repeat-x;"></td>
		<%
		next
		
		else

			for i=1 to dayarray(j)
				if RSUSER(i & "_" & j) <> "" then%>
				<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/r.png'); background-repeat: repeat-x;" onmouseup="popup('setup.asp?id=7&amp;part=3&amp;user=<%=RSUSER("ID")%>&amp;period=<%=i%>&amp;day=<%=j%>')"><%=RSUSER(i & "_" & j)%></td>
				<%
				else
				%>
				<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/g.png'); background-repeat: repeat-x;" onmouseup="popup('setup.asp?id=7&amp;part=3&amp;user=<%=RSUSER("ID")%>&amp;period=<%=i%>&amp;day=<%=j%>')">Free</td>
				<%
				end if
			next
		for i=1 to daytot
		%>
		<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/na_a.png'); background-repeat: repeat-x;"></td>
		<%
		next
		end if
		set daytot = nothing
		%>		
	</tr>
	<%
	next
	%>
</table>