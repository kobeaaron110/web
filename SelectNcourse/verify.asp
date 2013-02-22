<% 'If Session("LogonLevel")=2 or Session("LogonLevel")=3 then
   ' Response.Redirect "cexam.asp"
	if Session("LogonLevel")="ST" Then
		'if date<=#2012/09/05# then
		if date<=#2013/02/21# then
			Response.Redirect "index.asp"
		else
			Response.Redirect "index_print.asp"
		end if 
	elseif Session("LogonLevel")="Admin" Then
		Response.Redirect "SelectMan.asp"
	else
		Response.Redirect "login.asp"
	end if 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<body>
</body>
</html>
