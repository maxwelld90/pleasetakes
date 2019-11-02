// Server-ML.co.uk PleaseTakes Version 1
// Copyright (c) Server-ML.co.uk 2006
// Administrative Javascript File

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

function pjump()
	{
	if (document.add.p1.value.length == 3)
		{
		document.add.p2.focus();
		}
	}

function printpage(pid,page)
	{
	var browserName=navigator.appName;
		if(browserName == 'Microsoft Internet Explorer')
		{
		print();
		}
		else
		{
		location.href='print.asp?id=' + pid 
		}
	}

function popup(URL)
	{
	day = new Date();
	id = day.getTime();
	eval("page" + id + " = window.open(URL, '" + id + "', 'toolbar=0,scrollbars=1,location=0,statusbar=0,menubar=0,resizable=0,width=690,height=550,left = 20,top = 20');");
	}

function noalpha(contents)
	{
    if (((contents / contents) != 1) && (contents != 0)) {alert('Only Numbers Please!')}
	}

function colorRow(table, mouse)
	{
	table.style.backgroundColor = mouse==1?'#F7C077':'';
	}

function colorRowLight(table, mouse)
	{
	table.style.backgroundColor = mouse==1?'#FAD9AB':'';
	}

function showdetail(element) {
	var browserName=navigator.appName;
	if(document.getElementById('list_' + element ).style.display == 'none')
	{
		if(browserName == 'Microsoft Internet Explorer')
		{
		document.getElementById('list_' + element ).style.display='block';
		}
		else
		{
		document.getElementById('list_' + element ).style.display='table-row';
		}
	}
	else
	{
	document.getElementById('list_' + element ).style.display='none';
	}
	}