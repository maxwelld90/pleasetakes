// Server-ML.co.uk PleaseTakes Version 1
// Copyright (c) Server-ML.co.uk 2006
// Signup Screen Javascript File

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
	if (document.frm.p1.value.length == 3)
		{
		document.frm.p2.focus();
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