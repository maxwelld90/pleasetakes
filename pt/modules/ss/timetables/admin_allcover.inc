<%
	if GETDOW(date()) = "Sunday" then
		DOW = 1
	elseif GETDOW(date()) = "Monday" then
		DOW = 2
	elseif GETDOW(date()) = "Tuesday" then
		DOW = 3
	elseif GETDOW(date()) = "Wednesday" then
		DOW = 4
	elseif GETDOW(date()) = "Thursday" then
		DOW = 5
	elseif GETDOW(date()) = "Friday" then
		DOW = 6
	elseif GETDOW(date()) = "Saturday" then
		DOW = 7
	end if

RSMAXSQL = "SELECT MAX(Totals) AS Highest FROM Periods"

Set RSMAX = Server.CreateObject("Adodb.RecordSet")
RSMAX.Open RSMAXSQL, dataconn, adopenkeyset, adlockoptimistic

RSTOTABSENTSQL = "SELECT * FROM Attendance WHERE DAYDATE = #" & SQLDate(coverday) & "# AND DAY = " & daydow

Set RSTOTABSENT = Server.CreateObject("Adodb.RecordSet")
RSTOTABSENT.Open RSTOTABSENTSQL, dataconn, adopenkeyset, adlockoptimistic

RSDAYSQL = "SELECT * FROM Periods WHERE ID = " & DAYDOW

Set RSDAY = Server.CreateObject("Adodb.RecordSet")
RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic

if RSTOTABSENT.RECORDCOUNT = 0 then

response.redirect "cover.asp?err=1"

else
%>

<table class="m_l_tt" cellspacing="0" cellpadding="0">
	<tr>
		<td class="m_l_tt_td" style="font-size: 8pt;"><%=coverday%></td>
		<%
		for i=1 to RSMAX("Highest")
		%>
		<td class="m_l_tt_td" style="font-size: 13pt;"><%=i%></td>
		<%
		next
		%>
	</tr>

<%
do until RSTOTABSENT.EOF

	RSUSERSQL = "SELECT * FROM Timetables WHERE ID = " & RSTOTABSENT("USER")

	Set RSUSER = Server.CreateObject("Adodb.RecordSet")
	RSUSER.Open RSUSERSQL, dataconn, adopenkeyset, adlockoptimistic
	
%>
	<tr>
		<td class="m_l_tt_td" style="font-size: 8pt;" onmouseup="location.href='/pt/modules/ss/db/edit.asp?edittype=3&amp;type=2&amp;user=<%=RSUSER("ID")%>&amp;total=<%=RSDAY("Totals")%>&amp;whole=1&amp;day=<%=daydow%>&amp;coverday=<%=coverday%>'">
		<%=RSUSER("LN")%>, <%=left(RSUSER("FN"),1)%>.
		</td>
		<%
		daytot = RSMAX("Highest") - RSDAY("Totals")
		
		if daytot = RSMAX("Highest") then
			for i=1 to RSMAX("Highest")
		%>
		<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/na_a.png'); background-repeat: repeat-x;"></td>
		<%
		next
		
		else
		
			for i=1 to RSDAY("Totals")
				if RSUSER(i & "_" & DAYDOW) <> "" then
					if RSTOTABSENT(i & "_" & DAYDOW) <> "" then%>
				<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/r.png'); background-repeat: repeat-x; font-size: 12pt; font-weight: bold;" onmouseup="location.href='/pt/modules/ss/db/edit.asp?edittype=3&amp;type=2&amp;user=<%=RSUSER("ID")%>&amp;period=<%=i%>&amp;day=<%=daydow%>&amp;coverday=<%=coverday%>'">A</td>
				<%
					else
				%>
				<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/g.png'); background-repeat: repeat-x; font-size: 12pt; font-weight: bold;" onmouseup="location.href='/pt/modules/ss/db/edit.asp?edittype=3&amp;type=2&amp;user=<%=RSUSER("ID")%>&amp;period=<%=i%>&amp;day=<%=daydow%>&amp;coverday=<%=coverday%>'">P</td>
				<%
					end if				
				else
					if RSTOTABSENT(i & "_" & DAYDOW) <> "" then
				%>
				<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/nc.png'); background-repeat: repeat-x; font-size: 12pt; font-weight: bold;" onmouseup="location.href='/pt/modules/ss/db/edit.asp?edittype=3&amp;type=2&amp;user=<%=RSUSER("ID")%>&amp;period=<%=i%>&amp;day=<%=daydow%>&amp;coverday=<%=coverday%>'">A</td>
				<%
					else
				%>
				<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/nc.png'); background-repeat: repeat-x; font-size: 12pt; font-weight: bold;" onmouseup="location.href='/pt/modules/ss/db/edit.asp?edittype=3&amp;type=2&amp;user=<%=RSUSER("ID")%>&amp;period=<%=i%>&amp;day=<%=daydow%>&amp;coverday=<%=coverday%>'">P</td>
				<%
					end if
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
	RSUSER.close
	set RSUSER = nothing
	
RSTOTABSENT.MOVENEXT
loop
%>
</table>
<%
end if
%>