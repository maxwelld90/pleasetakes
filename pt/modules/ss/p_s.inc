<!--
Server-ML.co.uk PleaseTakes Version 1
Copyright (c) Server-ML.co.uk 2006

Server Output Begins
-->

<!--#include virtual="/pt/modules/ss/xml_settings_s.inc"-->
<!--#include virtual="/pt/modules/ss/date.time.inc"-->

<%
const adOpenForwardOnly = 0
const adOpenKeySet = 1
const adOpenDynamic = 2
const adOpenStatic = 3
const adLockReadOnly = 1
const adLockPessimistic = 2
const adLockOptimistic = 3
const adLockBatchOptimistic = 4

Set dataconn = server.createobject("adodb.connection")
dataopen = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='D:\Websites\Data\PleaseTakes\data.mdb'"
dataconn.Open dataopen

Set userconn = server.createobject("adodb.connection")
useropen = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='D:\Websites\Data\PleaseTakes\Users.mdb'"
userconn.Open useropen

Set backupconn = server.createobject("adodb.connection")
backupopen = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='D:\Websites\Data\PleaseTakes\Backup.mdb'"
backupconn.Open backupopen

function removeDomain (tempStr) 
	xStr = inStr(tempStr, "@")
	tempStr = left(tempStr, xStr-1)
	removeDomain = tempStr
end function
%>