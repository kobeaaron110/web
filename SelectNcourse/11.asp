<% Session.codepage=65001%>
<% Response.CharSet="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Language" content="zh-tw">
		<meta http-equiv="Content-Type" content="text/html; charset=big5" />
		<STYLE type="text/css">
		A {text-decoration:none}
		A:hover {color: #FF0000; text-decoration: underline}
		</STYLE>
		<title>選擇夜輔暨晚自習參加科目</title>
	</head>
	<body bgcolor="#C7EDCC">
<Br><Br>
<%
	
	Response.Write("<table border='0' width='100%' ID='Table1'>")
	Response.Write("<form method='POST' name='Select' action='11.asp' ID='Form1'>")
	Response.Write("<tr>")
	Response.Write("<td> 參加科目：<a href='My.rar' target='_blank'>下載 夜間輔導暨晚自習實施要點</td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	
    Response.Write("<tr>")
    Response.Write("<br><p align='center'><td width='50%' height='60'><input class='button' type='submit' value='下一步' name='nextsubmit' onclick='return checkvalue()' ID='Submit1'>")
    Response.Write("</td>")
    Response.Write("</tr>")
    
	Response.Write("</form>")
	Response.Write("</table>")
	Response.Write("</body>")
%> 
</html>
