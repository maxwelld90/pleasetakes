// Server-ML.co.uk PleaseTakes Version 1
// Copyright (c) Server-ML.co.uk 2006
// Login Screen Javascript File

var daymsg;
now = new Date

if (now.getHours() < 12)
	{
	daymsg = "Good Morning"
	}
else if (now.getHours() < 17)
	{
	daymsg = "Good Afternoon"
	}

else
	{
	daymsg = "Good Evening"
	}

function popup(URL)
	{
	day = new Date();
	id = day.getTime();
	eval("page" + id + " = window.open(URL, '" + id + "', 'toolbar=0,scrollbars=1,location=0,statusbar=0,menubar=0,resizable=0,width=690,height=550,left = 20,top = 20');");
	}

function load()
	{
	document.frm.un.focus();
	document.body.scroll="no";
	}

function offlineload()
	{
	document.body.scroll="no";
	}

function pjump()
	{
	if (document.frm.p1.value.length == 3)
		{
		document.frm.p2.focus();
		}
	}