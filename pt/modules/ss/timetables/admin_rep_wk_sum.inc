<%
DAYDOW = j
DAYDATE = DateAdd("d",j - 1,getweekstart(date()))

RSMAXSQL = "SELECT MAX(Totals) AS Highest FROM Periods"

Set RSMAX = Server.CreateObject("Adodb.RecordSet")
RSMAX.Open RSMAXSQL, dataconn, adopenkeyset, adlockoptimistic

RSTOTABSENTSQL = "SELECT * FROM Attendance WHERE DAY = " & DAYDOW & "AND DAYDATE = #" & SQLDate(DAYDATE) & "#"

Set RSTOTABSENT = Server.CreateObject("Adodb.RecordSet")
RSTOTABSENT.Open RSTOTABSENTSQL, dataconn, adopenkeyset, adlockoptimistic

RSDAYSQL = "SELECT * FROM Periods WHERE ID = " & DAYDOW

Set RSDAY = Server.CreateObject("Adodb.RecordSet")
RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic

if RSTOTABSENT.RECORDCOUNT = 0 then
%>
No Staff Are Listed As Being Absent For <%=getDOW(DAYDATE)%>, <%=DAYDATE%>!
<%
else
%>

<table class="m_l_tt" cellspacing="0" cellpadding="0">
	<tr>
		<td class="m_l_tt_td" style="font-size: 8pt;"><%=DAYDATE%></td>
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
		<td class="m_l_tt_td" style="font-size: 8pt;">
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
				RSCLASSSQL = "SELECT [" & i & "_" & DAYDOW & "] FROM Timetables WHERE ID = " & RSTOTABSENT("USER") 

				Set RSCLASS = Server.CreateObject("Adodb.RecordSet")
				RSCLASS.Open RSCLASSSQL, dataconn, adopenkeyset, adlockoptimistic
				
				RSCOVERNSQL = "SELECT * FROM Cover WHERE FOR = " & RSTOTABSENT("USER") & " AND DAY = " & DAYDOW & " AND DAYDATE = #" & SQLDate(DAYDATE) & "# AND PERIOD = " & i

				Set RSCOVERN = Server.CreateObject("Adodb.RecordSet")
				RSCOVERN.Open RSCOVERNSQL, dataconn, adopenkeyset, adlockoptimistic

				if RSUSER(i & "_" & DAYDOW) <> "" then
					if RSTOTABSENT(i & "_" & DAYDOW) <> "" then
					
					if (RSUSER(i & "_" & DAYDOW) <> "") and (RSTOTABSENT(i & "_" & DAYDOW) <> "") and (RSCOVERN.RECORDCOUNT = 0) then
				%>
				<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/r.png'); background-repeat: repeat-x; font-size: 22pt; font-weight: bold;">?</td>	
				<%
					else
				RSCOVERDETSQL = "SELECT FN, LN FROM Timetables WHERE ID = " & RSCOVERN("COVERING")

				Set RSCOVERDET = Server.CreateObject("Adodb.RecordSet")
				RSCOVERDET.Open RSCOVERDETSQL, dataconn, adopenkeyset, adlockoptimistic
						if RSCOVERDET.RECORDCOUNT = 0 then
							RSCOVERDET2SQL = "SELECT FN, LN FROM OCover WHERE ID = " & RSCOVERN("COVERING")

							Set RSCOVERDET2 = Server.CreateObject("Adodb.RecordSet")
							RSCOVERDET2.Open RSCOVERDET2SQL, dataconn, adopenkeyset, adlockoptimistic
				%>
				<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/r.png'); background-repeat: repeat-x; font-size: 7pt;">(<%=RSCLASS(i & "_" & DAYDOW)%>)<br><%=left(RSCOVERDET2("FN"),1)%>.&nbsp;<%=RSCOVERDET2("LN")%></td>
				<%
						else
				%>
				<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/r.png'); background-repeat: repeat-x; font-size: 7pt;">(<%=RSCLASS(i & "_" & DAYDOW)%>)<br><%=left(RSCOVERDET("FN"),1)%>.&nbsp;<%=RSCOVERDET("LN")%></td>
				<%
						end if
						end if
					else
				%>
				<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/g.png'); background-repeat: repeat-x; font-size: 12pt; font-weight: bold;">P</td>
				<%
					end if
				else
					if RSTOTABSENT(i & "_" & DAYDOW) <> "" then
				%>
				<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/nc.png'); background-repeat: repeat-x; font-size: 12pt; font-weight: bold;">NC</td>
				<%
					else
				%>
				<td class="m_l_tt_td" style="background-image: url('/pt/media/tt/nc.png'); background-repeat: repeat-x; font-size: 12pt; font-weight: bold;">NC</td>
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