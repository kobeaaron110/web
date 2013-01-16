<% 'If Session("LogonLevel")=2 or Session("LogonLevel")=3 then
   ' Response.Redirect "cexam.asp"
	if Session("LogonLevel")="ST" Then
		'if date<"2011/1/29" then
			Response.Redirect "index.asp"
		'else
		'	Response.Redirect "index_print.asp"
		'end if 
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
