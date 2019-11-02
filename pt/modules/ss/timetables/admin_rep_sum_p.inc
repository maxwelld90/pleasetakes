<%
RSMAXSQL = "SELECT MAX(Totals) AS Highest FROM Periods"

Set RSMAX = Server.CreateObject("Adodb.RecordSet")
RSMAX.Open RSMAXSQL, dataconn, adopenkeyset, adlockoptimistic

RSTOTABSENTSQL = "SELECT * FROM Attendance WHERE DAY = " & daydow & " AND DAYDATE = #" & SQLDate(daydate) & "#"

Set RSTOTABSENT = Server.CreateObject("Adodb.RecordSet")
RSTOTABSENT.Open RSTOTABSENTSQL, dataconn, adopenkeyset, adlockoptimistic

RSDAYSQL = "SELECT * FROM Periods WHERE ID = " & daydow

Set RSDAY = Server.CreateObject("Adodb.RecordSet")
RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic

if RSTOTABSENT.RECORDCOUNT = 0 then
%>
No Staff Are Listed As Being Absent For <%=getdow(daydow)%>!
<%
else
%>

<table class="m_l_tt_p" cellspacing="0" cellpadding="0">
	<tr>
		<td class="m_l_tt_td_p" style="font-size: 8pt;"><%=daydate%></td>
		<%
		for i=1 to RSMAX("Highest")
		%>
		<td class="m_l_tt_td_p" style="font-size: 18pt;"><%=i%></td>
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
		<td class="m_l_tt_td_p" style="font-size: 8pt;">
		<%=RSUSER("LN")%>, <%=left(RSUSER("FN"),1)%>.
		</td>
		<%
		daytot = RSMAX("Highest") - RSDAY("Totals")
		
		if daytot = RSMAX("Highest") then
			for i=1 to RSMAX("Highest")
		%>
		<td class="m_l_tt_td_p"></td>
		<%
		next
		
		else
		
			for i=1 to RSDAY("Totals")
				RSCLASSSQL = "SELECT [" & i & "_" & daydow & "] FROM Timetables WHERE ID = " & RSTOTABSENT("USER") 

				Set RSCLASS = Server.CreateObject("Adodb.RecordSet")
				RSCLASS.Open RSCLASSSQL, dataconn, adopenkeyset, adlockoptimistic
				
				RSCOVERNSQL = "SELECT * FROM Cover WHERE FOR = " & RSTOTABSENT("USER") & " AND DAY = " & daydow & " AND DAYDATE = #" & SQLDate(daydate) & "# AND PERIOD = " & i

				Set RSCOVERN = Server.CreateObject("Adodb.RecordSet")
				RSCOVERN.Open RSCOVERNSQL, dataconn, adopenkeyset, adlockoptimistic

				if RSUSER(i & "_" & daydow) <> "" then
					if RSTOTABSENT(i & "_" & daydow) <> "" then
					
					if (RSUSER(i & "_" & daydow) <> "") and (RSTOTABSENT(i & "_" & daydow) <> "") and (RSCOVERN.RECORDCOUNT = 0) then
				%>
				<td class="m_l_tt_td_p" style="font-size: 22pt; font-weight: bold;">?</td>	
				<%
					else

				RSCOVERDETSQL = "SELECT FN, LN FROM Timetables WHERE ID = " & RSCOVERN("COVERING")

				Set RSCOVERDET = Server.CreateObject("Adodb.RecordSet")
				RSCOVERDET.Open RSCOVERDETSQL, dataconn, adopenkeyset, adlockoptimistic
				
						if (rscovern("ocover") = "1") then
							RSCOVERDET2SQL = "SELECT FN, LN FROM OCover WHERE ID = " & RSCOVERN("COVERING")
			
							Set RSCOVERDET2 = Server.CreateObject("Adodb.RecordSet")
							RSCOVERDET2.Open RSCOVERDET2SQL, dataconn, adopenkeyset, adlockoptimistic
				%>
				<td class="m_l_tt_td_p" style="font-size: 7pt;">OUT:(<%=RSCLASS(i & "_" & daydow)%>)<br><%=left(RSCOVERDET2("FN"),1)%>.&nbsp;<%if (len(RSCOVERDET2("LN")) > 12) then%><%=left(RSCOVERDET2("LN"),8)%>...<%else%><%=RSCOVERDET2("LN")%><%end if%></td>
				<%
						else
				%>
				<td class="m_l_tt_td_p" style="font-size: 7pt;">(<%=RSCLASS(i & "_" & daydow)%>)<br><%=left(RSCOVERDET("FN"),1)%>.&nbsp;<%if (len(RSCOVERDET("LN")) > 12) then%><%=left(RSCOVERDET("LN"),8)%>...<%else%><%=RSCOVERDET("LN")%><%end if%></td>
				<%
						end if
						end if
					else
				%>
				<td class="m_l_tt_td_p" style="font-size: 12pt; font-weight: bold;">P</td>
				<%
					end if
				else
					if RSTOTABSENT(i & "_" & daydow) <> "" then
				%>
				<td class="m_l_tt_td_p" style="font-size: 12pt; font-weight: bold;">NC</td>
				<%
					else
				%>
				<td class="m_l_tt_td_p" style="font-size: 12pt; font-weight: bold;">NC</td>
				<%
					end if
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
	RSUSER.close
	set RSUSER = nothing
	
RSTOTABSENT.MOVENEXT
loop
%>
</table>
<%
end if
%>