// Server-ML.co.uk PleaseTakes Version 1
// Copyright (c) Server-ML.co.uk 2006
// Standard Account Javascript File

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

function colorRow(table, mouse)
	{
	table.style.backgroundColor = mouse==1?'#80A0EE':'';
	}
function colorRowLight(table, mouse)
	{
	table.style.backgroundColor = mouse==1?'#D1E7FC':'';
	}