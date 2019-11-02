<%
if var_est_enabled_weekends = "1" then
	if DOW = "7" then
		if date_hour => 15 then
		else
			response.redirect "/pt/admin/backup.asp?id=1&err=1"
		end if
	else
		response.redirect "/pt/admin/backup.asp?id=1&err=1"
	end if
else
	if DOW = "6" then
		if date_hour => 15 then
		else
			response.redirect "/pt/admin/backup.asp?id=1&err=1"
		end if
	else
		response.redirect "/pt/admin/backup.asp?id=1&err=1"
	end if
end if
%>				