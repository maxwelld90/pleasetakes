<div class="m_r_ti">My PleaseTakes For Today</div>
<div class="m_r_t">
<!--#include virtual="/pt/modules/ss/timetables/std_rbar_sum.inc"-->
</div>


<%
RSMAXSQL = "SELECT MAX(Totals) AS Highest FROM Periods"

Set RSMAX = Server.CreateObject("Adodb.RecordSet")
RSMAX.Open RSMAXSQL, dataconn, adopenkeyset, adlockoptimistic

for i = 1 to RSMAX("HIGHEST")

	RSCOVERINGSQL = "SELECT * FROM COVER WHERE COVERING = " & session("sess_ttid") & " AND DAY = " & DOW & " AND PERIOD = " & i

	Set RSCOVERING = Server.CreateObject("Adodb.RecordSet")
	RSCOVERING.Open RSCOVERINGSQL, dataconn, adopenkeyset, adlockoptimistic
	
	if RSCOVERING.RECORDCOUNT = 0 then

		RSTEACHINGSQL = "SELECT * FROM Timetables WHERE ID = " & session("sess_ttid")

		Set RSTEACHING = Server.CreateObject("Adodb.RecordSet")
		RSTEACHING.Open RSTEACHINGSQL, dataconn, adopenkeyset, adlockoptimistic
		
		RSTEACHINGSQL = "SELECT * FROM Timetables WHERE ID = " & session("sess_ttid")

		Set RSTEACHING = Server.CreateObject("Adodb.RecordSet")
		RSTEACHING.Open RSTEACHINGSQL, dataconn, adopenkeyset, adlockoptimistic
		
		if RSTEACHING(i & "_" & DOW) <> "" then
%>
<div class="m_r_ti">Period <%=i%> - Teaching</div>
<div class="m_r_t">
Class <b><%=RSTEACHING(i & "_" & DOW)%></b><br>
Room <b><%=RSTEACHING("R" & i & "_" & DOW)%></b>
</div>
<%
		else
		end if
	else
	
	RSCOVERINGINFOSQL = "SELECT * FROM Timetables WHERE ID = " & RSCOVERING("FOR")

	Set RSCOVERINGINFO = Server.CreateObject("Adodb.RecordSet")
	RSCOVERINGINFO.Open RSCOVERINGINFOSQL, dataconn, adopenkeyset, adlockoptimistic
%>
<div class="m_r_ti">Period <%=i%> - Covering</div>
<div class="m_r_t">
Class <b><%=RSCOVERINGINFO(i & "_" & DOW)%></b><br>
Room <b><%=RSCOVERINGINFO("R" & i & "_" & DOW)%></b><br>
For <b><%=left(RSCOVERINGINFO("FN"),1)%>.&nbsp;<%=RSCOVERINGINFO("LN")%></b>
</div>
<%
	end if
next
%>