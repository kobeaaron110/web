<% Session.codepage=65001%>
<% Response.CharSet="utf-8" %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�]���[�ߦ۲�</title>
<base target="main">
</head>

<body bgcolor="#C7EDCC">

<Br><Br>

<form method="POST" autocomplete="off" name="menu" action="menu.asp">
<table border="1" width="100%" cellspacing="0" height="1" cellpadding="3">
<%
Response.Write("<tr>")
Response.Write("<td width=100% height=27>")
Response.Write("<p align=center><font size=4>�]���[�ߦ۲�</font></td>")
Response.Write("</tr>")
if date<#2014/2/21# then
	Response.Write("<tr>")
	Response.Write("<td width=100% height=23><font SIZE=2><a href='main.asp' target='mainFrame'> ��ܩ]���[�ߦ۲߰ѥ[��� </a></font></td>")
	Response.Write("</tr>")
end if
	Response.Write("<tr>")
	Response.Write("<td width=100% height=24><font SIZE=2><a href='print.asp' target='mainFrame'> �C�L�P�N�� </a></font></td>")
	Response.Write("</tr>")
Response.Write("</table>")

Response.Write("<Br><Br>")

Response.Write("<p align=center><a href='login.asp?logout=yes' target='_top'><font SIZE=6 face=�з��� color=Blue><strong>�n�X</strong></font></a>")
Response.Write("</form>")
%>
      
<Br><Br>

<font SIZE="2"></font>   
      
</body>      
      
</html>      
