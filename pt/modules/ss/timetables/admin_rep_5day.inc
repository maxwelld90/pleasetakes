<%
if (request("order")) = "DEPT" then
RSUSERSQL = "SELECT * FROM Timetables ORDER BY DEPT, LN"
elseif (request("order")) = "TI" then
RSUSERSQL = "SELECT * FROM Timetables ORDER BY CATEGORY, LN"
else
RSUSERSQL = "SELECT * FROM Timetables ORDER BY LN"
end if
	
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

Dim dayarray2(6)

for i=1 to 6
dayarray2(i) = RSDAY("TOTALS")
RSDAY.MOVENEXT
next
RSDAY.MOVEFIRST

do until RSUSER.EOF
%>
			<span style="font-size: 12pt; font-weight: bold;"><%=RSUSER("LN")%>, <%=left(RSUSER("FN"),1)%>.</span>&nbsp;
					<%
					if RSUSER("CATEGORY") = "T" then
					response.write "Teacher"
					elseif RSUSER("CATEGORY") = "PT" then
					response.write "Principal Teacher"
					elseif RSUSER("CATEGORY") = "DHT" then
					response.write "Deputy Head Teacher"
					elseif RSUSER("CATEGORY") = "HT" then
					response.write "Head Teacher"
					elseif RSUSER("CATEGORY") = "PS" then
					response.write "Pupil Support"
					elseif RSUSER("CATEGORY") = "OC" then
					response.write "Outside Cover"
					else
					response.write "Unknown Position"
					end if
					%>,

					<%
RSDEPTSQL = "SELECT * FROM DEPARTMENTS WHERE DEPTID = " & RSUSER("DEPT")

Set RSDEPT = Server.CreateObject("Adodb.RecordSet")
RSDEPT.Open RSDEPTSQL, dataconn, adopenkeyset, adlockoptimistic
%>
					<%=RSDEPT("FULL")%>
			<hr size="1">
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
		<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/na_a.png'); background-repeat: repeat-x;"></td>
		<%
		next
		
		else

			for i=1 to dayarray2(j)
				if RSUSER(i & "_" & j) <> "" then%>
				<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/r.png'); background-repeat: repeat-x;" onmouseup="popup('popup.asp?id=1&amp;user=<%=RSUSER("ID")%>&amp;period=<%=i%>&amp;day=<%=j%>')"><%=RSUSER(i & "_" & j)%></td>
				<%
				else
				%>
				<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/g.png'); background-repeat: repeat-x;" onmouseup="popup('popup.asp?id=1&amp;user=<%=RSUSER("ID")%>&amp;period=<%=i%>&amp;day=<%=j%>')">Free</td>
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
<div class="m_l_sel_sep"></div>
<%
	RSDAY.MOVEFIRST
RSUSER.MOVENEXT
loop

RSDEPT.close
RSMAX.close
RSDAY.close
RSUSER.close

set RSDEPT = nothing
set RSMAX = nothing
set RSDAY = nothing
set RSUSER = nothing
%>