<% if Session("UsrLevel")=99 then
	 if Session("TableType1_2")=1 then
		Response.Redirect "admin.asp"
	 else
		Response.Redirect "login2.asp?logout2=yes"
	 end if 
   else
     Response.Redirect "login2.asp"
   end if 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<body>
</body>
</html>
