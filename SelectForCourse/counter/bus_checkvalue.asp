<script language="javascript"> 
function checkvalue()
{
	var thisForm = document.Select;
	var s = thisForm.Select1;
    if (s.selectedIndex==0)
    {
		//alert( "BusStop is blank , Please Select!!" );
		//return false;
    }
}

function eng_basic()
{
	if (document.Select.Subject3.checked)
		document.Select.Subject4.disabled = true;
	else
		document.Select.Subject4.disabled = false;
}
function eng_adv()
{
	if (document.Select.Subject4.checked)
		document.Select.Subject3.disabled = true;
	else
		document.Select.Subject3.disabled = false;
}

function pic_basic()
{
	if (document.Select.Subject5.checked)
		document.Select.Subject6.disabled = true;
	else
		document.Select.Subject6.disabled = false;
}
function pic_adv()
{
	if (document.Select.Subject6.checked)
		document.Select.Subject5.disabled = true;
	else
		document.Select.Subject5.disabled = false;
}

function Bus_Stop0()
{
	var thisForm = document.Select;
	var s = thisForm.Select1;
	var i;	
	for ( i= s.length-1 ; i > 0 ; i-- )
	{
	s.remove( i );
	}
}
function Bus_Stop1()
{
	var thisForm = document.Select;
	var s = document.getElementById('Select1');
	var t;
	var v;
	var i;
	
	for ( i= s.length-1 ; i > 0 ; i-- )
	{
	s.remove( i );
	}
	
	<%
	set rs1=conn.Execute("Select * from 站名資料 Where 車號=1 order by 站序")
	while not rs1.eof
		response.write "t='" & rs1("代碼") & " " & rs1("站名") & "';" & chr(13)
		response.write "v='" & rs1("代碼") & "';" & chr(13)
	%>
	var new_option = new Option(t,v);
	s.options.add(new_option);
	<%
		rs1.movenext
	wend
	%>
	//s.options.add(new_option);
}
function Bus_Stop6()
{
	var thisForm = document.Select;
	var s = document.getElementById('Select1');
	var t;
	var v;
	var i;	
	for ( i= s.length-1 ; i > 0 ; i-- )
	{
	s.remove( i );
	}
	
	<%
	set rs1=conn.Execute("Select * from 站名資料 Where 車號=6 order by 站序")
	while not rs1.eof
		response.write "t='" & rs1("代碼") & " " & rs1("站名") & "';" & chr(13)
		response.write "v='" & rs1("代碼") & "';" & chr(13)
	%>
	var new_option = new Option(t,v);
	s.options.add(new_option);
	<%
		rs1.movenext
	wend
	%>
}
function Bus_Stop9()
{
	var thisForm = document.Select;
	var s = document.getElementById('Select1');
	var t;
	var v;
	var i;	
	for ( i= s.length-1 ; i > 0 ; i-- )
	{
	s.remove( i );
	}
	
	<%
	set rs1=conn.Execute("Select * from 站名資料 Where 車號=9 order by 站序")
	while not rs1.eof
		response.write "t='" & rs1("代碼") & " " & rs1("站名") & "';" & chr(13)
		response.write "v='" & rs1("代碼") & "';" & chr(13)
	%>
	var new_option = new Option(t,v);
	s.options.add(new_option);
	<%
		rs1.movenext
	wend
	%>
}
function Bus_Stop15()
{
	var thisForm = document.Select;
	var s = document.getElementById('Select1');
	var t;
	var v;
	var i;	
	for ( i= s.length-1 ; i > 0 ; i-- )
	{
	s.remove( i );
	}
	
	<%
	set rs1=conn.Execute("Select * from 站名資料 Where 車號=15 order by 站序")
	while not rs1.eof
		response.write "t='" & rs1("代碼") & " " & rs1("站名") & "';" & chr(13)
		response.write "v='" & rs1("代碼") & "';" & chr(13)
	%>
	var new_option = new Option(t,v);
	s.options.add(new_option);
	<%
		rs1.movenext
	wend
	%>
}
function Bus_Stop16()
{
	var thisForm = document.Select;
	var s = document.getElementById('Select1');
	var t;
	var v;
	var i;	
	for ( i= s.length-1 ; i > 0 ; i-- )
	{
	s.remove( i );
	}
	
	<%
	set rs1=conn.Execute("Select * from 站名資料 Where 車號=16 order by 站序")
	while not rs1.eof
		response.write "t='" & rs1("代碼") & " " & rs1("站名") & "';" & chr(13)
		response.write "v='" & rs1("代碼") & "';" & chr(13)
	%>
	var new_option = new Option(t,v);
	s.options.add(new_option);
	<%
		rs1.movenext
	wend
	%>
}
function Bus_Stop20()
{
	var thisForm = document.Select;
	var s = document.getElementById('Select1');
	var t;
	var v;
	var i;	
	for ( i= s.length-1 ; i > 0 ; i-- )
	{
	s.remove( i );
	}
	
	<%
	set rs1=conn.Execute("Select * from 站名資料 Where 車號=20 order by 站序")
	while not rs1.eof
		response.write "t='" & rs1("代碼") & " " & rs1("站名") & "';" & chr(13)
		response.write "v='" & rs1("代碼") & "';" & chr(13)
	%>
	var new_option = new Option(t,v);
	s.options.add(new_option);
	<%
		rs1.movenext
	wend
	%>
}
function Bus_Stop26()
{
	var thisForm = document.Select;
	var s = document.getElementById('Select1');
	var t;
	var v;
	var i;	
	for ( i= s.length-1 ; i > 0 ; i-- )
	{
	s.remove( i );
	}
	
	<%
	set rs1=conn.Execute("Select * from 站名資料 Where 車號=26 order by 站序")
	while not rs1.eof
		response.write "t='" & rs1("代碼") & " " & rs1("站名") & "';" & chr(13)
		response.write "v='" & rs1("代碼") & "';" & chr(13)
	%>
	var new_option = new Option(t,v);
	s.options.add(new_option);
	<%
		rs1.movenext
	wend
	%>
}
function Bus_Stop29()
{
	var thisForm = document.Select;
	var s = document.getElementById('Select1');
	var t;
	var v;
	var i;	
	for ( i= s.length-1 ; i > 0 ; i-- )
	{
	s.remove( i );
	}
	
	<%
	set rs1=conn.Execute("Select * from 站名資料 Where 車號=29 order by 站序")
	while not rs1.eof
		response.write "t='" & rs1("代碼") & " " & rs1("站名") & "';" & chr(13)
		response.write "v='" & rs1("代碼") & "';" & chr(13)
	%>
	var new_option = new Option(t,v);
	s.options.add(new_option);
	<%
		rs1.movenext
	wend
	%>
}
function Bus_Stop30()
{
	var thisForm = document.Select;
	var s = document.getElementById('Select1');
	var t;
	var v;
	var i;	
	for ( i= s.length-1 ; i > 0 ; i-- )
	{
	s.remove( i );
	}
	
	<%
	set rs1=conn.Execute("Select * from 站名資料 Where 車號=30 order by 站序")
	while not rs1.eof
		response.write "t='" & rs1("代碼") & " " & rs1("站名") & "';" & chr(13)
		response.write "v='" & rs1("代碼") & "';" & chr(13)
	%>
	var new_option = new Option(t,v);
	s.options.add(new_option);
	<%
		rs1.movenext
	wend
	%>
}
function Bus_Stop33()
{
	var thisForm = document.Select;
	var s = document.getElementById('Select1');
	var t;
	var v;
	var i;	
	for ( i= s.length-1 ; i > 0 ; i-- )
	{
	s.remove( i );
	}
	
	<%
	set rs1=conn.Execute("Select * from 站名資料 Where 車號=33 order by 站序")
	while not rs1.eof
		response.write "t='" & rs1("代碼") & " " & rs1("站名") & "';" & chr(13)
		response.write "v='" & rs1("代碼") & "';" & chr(13)
	%>
	var new_option = new Option(t,v);
	s.options.add(new_option);
	<%
		rs1.movenext
	wend
	%>
}
function Bus_Stop36()
{
	var thisForm = document.Select;
	var s = document.getElementById('Select1');
	var t;
	var v;
	var i;	
	for ( i= s.length-1 ; i > 0 ; i-- )
	{
	s.remove( i );
	}
	
	<%
	set rs1=conn.Execute("Select * from 站名資料 Where 車號=36 order by 站序")
	while not rs1.eof
		response.write "t='" & rs1("代碼") & " " & rs1("站名") & "';" & chr(13)
		response.write "v='" & rs1("代碼") & "';" & chr(13)
	%>
	var new_option = new Option(t,v);
	s.options.add(new_option);
	<%
		rs1.movenext
	wend
	%>
}
function Bus_Stop39()
{
	var thisForm = document.Select;
	var s = document.getElementById('Select1');
	var t;
	var v;
	var i;	
	for ( i= s.length-1 ; i > 0 ; i-- )
	{
	s.remove( i );
	}
	
	<%
	set rs1=conn.Execute("Select * from 站名資料 Where 車號=39 order by 站序")
	while not rs1.eof
		response.write "t='" & rs1("代碼") & " " & rs1("站名") & "';" & chr(13)
		response.write "v='" & rs1("代碼") & "';" & chr(13)
	%>
	var new_option = new Option(t,v);
	s.options.add(new_option);
	<%
		rs1.movenext
	wend
	%>
}
function Bus_Stop39B()
{
	var thisForm = document.Select;
	var s = document.getElementById('Select1');
	var t;
	var v;
	var i;	
	for ( i= s.length-1 ; i > 0 ; i-- )
	{
	s.remove( i );
	}
	
	<%
	set rs1=conn.Execute("Select * from 站名資料 Where 車號=79 order by 站序")
	while not rs1.eof
		response.write "t='" & rs1("代碼") & " " & rs1("站名") & "';" & chr(13)
		response.write "v='" & rs1("代碼") & "';" & chr(13)
	%>
	var new_option = new Option(t,v);
	s.options.add(new_option);
	<%
		rs1.movenext
	wend
	%>
}
function Bus_Stop42()
{
	var thisForm = document.Select;
	var s = document.getElementById('Select1');
	var t;
	var v;
	var i;	
	for ( i= s.length-1 ; i > 0 ; i-- )
	{
	s.remove( i );
	}
	
	<%
	set rs1=conn.Execute("Select * from 站名資料 Where 車號=42 order by 站序")
	while not rs1.eof
		response.write "t='" & rs1("代碼") & " " & rs1("站名") & "';" & chr(13)
		response.write "v='" & rs1("代碼") & "';" & chr(13)
	%>
	var new_option = new Option(t,v);
	s.options.add(new_option);
	<%
		rs1.movenext
	wend
	%>
}
function Bus_Stop88()
{
	var thisForm = document.Select;
	var s = document.getElementById('Select1');
	var t;
	var v;
	var i;	
	for ( i= s.length-1 ; i > 0 ; i-- )
	{
	s.remove( i );
	}
	
	<%
	set rs1=conn.Execute("Select * from 站名資料 Where 車號=80 order by 站序")
	while not rs1.eof
		response.write "t='" & rs1("站名") & "';" & chr(13)
		response.write "v='" & rs1("代碼") & "';" & chr(13)
	%>
	var new_option = new Option(t,v);
	s.options.add(new_option);
	<%
		rs1.movenext
	wend
	%>
}
</script> 
