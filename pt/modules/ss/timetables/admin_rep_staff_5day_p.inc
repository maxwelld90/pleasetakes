<%
RSUSERSQL = "SELECT * FROM Timetables WHERE ID = " & RSCHECK("ID")
	
Set RSUSER = Server.CreateObject("Adodb.RecordSet")
RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic

'Find Out Maximum Number Of Periods

RSMAXSQL = "SELECT MAX(Totals) AS Highest FROM Periods"

Set RSMAX = Server.CreateObject("Adodb.RecordSet")
RSMAX.Open RSMAXSQL, dataconn, adopenkeyset, adlockoptimistic

RSDAYSQL = "SELECT * FROM Periods"

Set RSDAY = Server.CreateObject("Adodb.RecordSet")
RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic

redim dayarray2(6)

for i=1 to 6
dayarray2(i) = RSDAY("TOTALS")
RSDAY.MOVENEXT
next
RSDAY.MOVEFIRST
%>
<table class="m_l_tt" style="border-color: #000;" cellspacing="0" cellpadding="0">
	<tr>
		<td class="m_l_tt_td_p"></td>
		<%
		for i=1 to RSMAX("Highest")
		%>
		<td class="m_l_tt_td_p" style="font-size: 13pt;"><%=i%></td>
		<%
		next
		%>
	</tr>

	<%
	RSDAY.MOVENEXT
	for j = 2 to 6
	%>
	<tr>
		<td class="m_l_tt_td_p" style="font-size: 13pt;">
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
		<td class="m_l_tt_td_p"></td>
		<%
		next
		
		else

			for i=1 to dayarray2(j)
				if RSUSER(i & "_" & j) <> "" then%>
				<td class="m_l_tt_td_p"><%=RSUSER(i & "_" & j)%></td>
				<%
				else
				%>
				<td class="m_l_tt_td_p">Free</td>
				<%
				end if
			next
		for i=1 to daytot
		%>
		<td class="m_l_tt_td_p"></td>
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
