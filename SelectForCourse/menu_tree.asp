<html>
<head>
	<meta http-equiv="content-type" content="text/html; charset=big5">
	
	<link href="common/link.css" rel="stylesheet" type="text/css">
	<script language="JavaScript" src="common/js06-06.js"></script>
</head>
<body background="images/bg_img.gif" bgcolor="#c8c8c8" text="#000000">
 
<SCRIPT Language="JavaScript">
foldersTree = gFld("", "index.asp")
insDoc(foldersTree, gLnk(1, "�]���ѥ[���", "select_sbj.asp"))
insDoc(foldersTree, gLnk(0, "�C�L�P�N��", "print.asp"))
//insDoc(foldersTree, gLnk(1, "�u�W���W�t��", "login.asp"))
//insDoc(foldersTree, gLnk(0, "��ƤU��", "��ƤU��"))
//insDoc(foldersTree, gLnk(0, "�d����", "Board/board.asp"))
initializeDocument()
</SCRIPT> 

&nbsp;<br> 
<p align="center"><font size="3" color="#000000">�H����ơG<b></font><font color="FE5AFB"><%=count%> </font></b><font size="3" color="#000000">�H </p>
<%
dim countStr(6)
	countPic=""
	For i = 1 to len(count)
		countStr(i) = mid(count,i,1)
		countPic=countPic+"<img src='images/digit/" & countStr(i) & ".gif' wide=16 height=20>"
	Next
	countPic="<p align='center'>�s���H��" & countPic & "</p>"
	response.write countPic
if Request.Cookies("passed")="newspassed" then
	Response.Cookies("passed")="newspassed"
	Response.Cookies("auth_num")=Request.Cookies("auth_num")
else
	Response.Cookies("passed")=""
	Response.Cookies("auth_num")=1
end if
%>
&nbsp;<br> 
<p align="center">��ĳ�ѪR�סG<br> 
	<font color="Red">1024 x 768</font></p>
</body> 
</html>

