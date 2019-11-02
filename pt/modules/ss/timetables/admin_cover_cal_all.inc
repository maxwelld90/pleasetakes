<!--
Server-ML.co.uk Calendar
For Use In PleaseTakes v1
-->

<%
datenow = date()
if request("date") = "" then
	thisMonth = request("month")
	thisYear = request("year")
		if thisMonth = "" then
			thisMonth = Month(Date)
		end if
		if thisYear = "" then
			thisYear = Year(Date)
		end if
	lastMonth = thisMonth - 1
	lastYear = thisYear
	nextMonth = thisMonth + 1
	nextYear = thisYear
		if lastMonth < 1 then
			lastMonth = 12
			lastYear = lastYear - 1
		end if
		if nextMonth >12 then
			nextMonth = 1
			nextYear = nextYear + 1
		end if

firstWeekDay = Weekday(CDate(thisMonth & "/15/" & thisYear))

leapTestNumbers = thisyear / 4
leapTest = leapTestNumbers - Round(leapTestNumbers)

if thismonth = 2 then
	if leapTest <> 0 then
		maxDays = 28
	else
		maxDays = 29
	end if
elseif (thisMonth = 4 or thisMonth = 6 or thisMonth = 9 or thisMonth = 11) then
	maxDays = 30
else
	maxDays = 31
end if

dayCounter = 1
weekDayCounter = 1
weekCounter = 1
%>

<table style="width: 100%;" cellspacing="0" cellpadding="0">
	<tr style="height: 35px; background: url('/pt/media/tt/cal_t.png') repeat-x;">
		<td>
			<table style="width: 100%;" cellspacing="0" cellpadding="0">
				<tr>
					<td style="width: 33%; padding-left: 5px; font-size: 10pt; font-weight: bold; letter-spacing: -1px;"><a href="cover.asp?id=6&amp;type=2&amp;month=<%=lastmonth%>&amp;year=<%=lastyear%>">&lt;&lt;<%=Cal_GetMonth(lastmonth)%></a></td>
					<td style="width: 33%; text-align: center; font-size: 14pt; font-weight: bold; letter-spacing: -2px;"><%=Cal_GetMonth(thismonth)%>, <%=thisyear%></td>
					<td style="width: 33%; padding-right: 5px; text-align: right; font-size: 10pt; font-weight: bold; letter-spacing: -1px;"><a href="cover.asp?id=6&amp;type=2&amp;month=<%=nextmonth%>&amp;year=<%=nextyear%>"><%=Cal_GetMonth(nextmonth)%>&gt;&gt;</a></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
			<table style="width: 100%;" cellspacing="0" cellpadding="0">
				<tr>
					<td>
						<table style="border-collapse: collapse; width: 100%;" cellspacing="0" cellpadding="0">
							<tr style="height: 45px;">
								<%
								i = 1
								do while i <= 7
								
									RSDAYSQL = "SELECT * FROM Periods WHERE DAYID = " & i

									Set RSDAY = Server.CreateObject("Adodb.RecordSet")
									RSDAY.Open RSDAYSQL, dataconn, adopenkeyset, adlockoptimistic
								%>
								<td class="m_l_tt_td" style="width: 78px; font-size: 12pt; font-weight: bold;"><%=RSDAY("DAYNAME")%></td>
								<%
									RSDAY.close
									set RSDAY = nothing

								i = i + 1
								loop
								%>
							</tr>
							<%
							do while maxDays >= dayCounter
							%>
							<tr style="height: 45px;">
							<%
							do while weekDayCounter <= 7
								if (weekCounter = 1 and weekDayCounter < firstWeekDay) or (dayCounter > maxDays) then
							%>
								<td class="m_l_tt_td" style="background: url('/pt/media/tt/na_a.png') repeat-x;"></td>
							<%
							else
								dateSelect = dayCounter & "/" & thisMonth & "/" & thisYear
								convdate = cdate(dateselect)
								
								if convdate < datenow then
								%>
									<td class="m_l_tt_td" style="width: 78px; font-size: 12pt; font-weight: bold; letter-spacing: -2px; background: url('/pt/media/tt/na.png') repeat-x;"><%=daycounter%></td>
								<%
								else
									if var_est_enabled_weekends <> "1" then
										if weekDayCounter = 1 then
									%>
									<td class="m_l_tt_td" style="width: 78px; font-size: 12pt; font-weight: bold; letter-spacing: -2px;<%if cdate(dateselect) = date() then%> background: url('/pt/media/tt/g.png') repeat-x;<%else%>  background: url('/pt/media/tt/nc.png') repeat-x;<%end if%>"><%=daycounter%></td>
									<%
										elseif weekDayCounter = 7 then
									%>
									<td class="m_l_tt_td" style="width: 78px; font-size: 12pt; font-weight: bold; letter-spacing: -2px;<%if cdate(dateselect) = date() then%> background: url('/pt/media/tt/g.png') repeat-x;<%else%>  background: url('/pt/media/tt/nc.png') repeat-x;<%end if%>"><%=daycounter%></td>
									<%
										else
									%>
									<td class="m_l_tt_td" style="width: 78px; font-size: 12pt; font-weight: bold; letter-spacing: -2px;<%if cdate(dateselect) = date() then%> background: url('/pt/media/tt/g.png') repeat-x;<%else%>  background: url('/pt/media/tt/nc.png') repeat-x;<%end if%>"><a href="cover.asp?id=2&amp;type=2&amp;coverday=<%=dateselect%>&amp;dow=<%=weekdaycounter%>"><%=daycounter%></a></td>
									<%
										end if
									else
									%>
									<td class="m_l_tt_td" style="width: 78px; font-size: 12pt; font-weight: bold; letter-spacing: -2px;<%if cdate(dateselect) = date() then%> background: url('/pt/media/tt/g.png') repeat-x;<%else%>  background: url('/pt/media/tt/nc.png') repeat-x;<%end if%>"><a href="cover.asp?id=2&amp;type=2&amp;coverday=<%=dateselect%>&amp;dow=<%=weekdaycounter%>"><%=daycounter%></a></td>
									<%
									end if

								end if
								dayCounter = dayCounter + 1
							end if
							
							weekDayCounter = weekDayCounter + 1
							loop
							weekCounter = weekCounter + 1
							
							if weekDayCounter > 7 then
								weekDayCounter = 1
							end if
							%>
							</tr>
							<%
							loop
							%>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%
end if
%>