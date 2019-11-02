<%
Set settingsXML = Server.CreateObject("Microsoft.XMLDOM")

settingsXML.async = False

settingsXML.Load (Server.MapPath("/pt/modules/xml/settings.xml"))

var_ptitle = settingsXML.documentElement.childNodes.item(0).childNodes.item(0).text
var_pname = settingsXML.documentElement.childNodes.item(0).childNodes.item(1).text
var_ver = settingsXML.documentElement.childNodes.item(0).childNodes.item(2).text
var_tz = settingsXML.documentElement.childNodes.item(0).childNodes.item(5).text
var_backuptime = settingsXML.documentElement.childNodes.item(0).childNodes.item(7).getAttribute("time")
var_backuprecent = settingsXML.documentElement.childNodes.item(0).childNodes.item(7).getAttribute("mostrecent")
var_maxperiods = settingsXML.documentElement.childNodes.item(0).childNodes.item(8).getAttribute("maximum")
var_emaildomain1 = settingsXML.documentElement.childNodes.item(0).childNodes.item(9).getAttribute("domain1")
var_usernames_std = settingsXML.documentElement.childNodes.item(0).childNodes.item(6).getAttribute("std")
var_usernames_admin = settingsXML.documentElement.childNodes.item(0).childNodes.item(6).getAttribute("admin")
var_firstlogin = settingsXML.documentElement.childNodes.item(0).childNodes.item(4).getAttribute("firstlogin")
var_est_enabled = settingsXML.documentElement.childNodes.item(0).childNodes.item(4).getAttribute("entire")
var_est_enabled_signup = settingsXML.documentElement.childNodes.item(0).childNodes.item(4).getAttribute("signup")
var_est_enabled_pin = settingsXML.documentElement.childNodes.item(0).childNodes.item(4).getAttribute("pin")
var_est_enabled_weekends = settingsXML.documentElement.childNodes.item(0).childNodes.item(4).getAttribute("weekends")
var_est_full = settingsXML.documentElement.childNodes.item(0).childNodes.item(3).text
var_est_short = settingsXML.documentElement.childNodes.item(0).childNodes.item(3).getAttribute("short")
var_est_logimg = settingsXML.documentElement.childNodes.item(0).childNodes.item(3).getAttribute("logimg")
var_est_stdimg = settingsXML.documentElement.childNodes.item(0).childNodes.item(3).getAttribute("stdimg")
%>
