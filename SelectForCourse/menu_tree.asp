<html>
<head>
	<meta http-equiv="content-type" content="text/html; charset=big5">
	
	<link href="common/link.css" rel="stylesheet" type="text/css">
	<script language="JavaScript" src="common/js06-06.js"></script>
</head>
<body background="images/bg_img.gif" bgcolor="#c8c8c8" text="#000000">
 
<SCRIPT Language="JavaScript">
foldersTree = gFld("", "index.asp")
insDoc(foldersTree, gLnk(1, "夜輔參加科目", "select_sbj.asp"))
insDoc(foldersTree, gLnk(0, "列印同意書", "print.asp"))
//insDoc(foldersTree, gLnk(1, "線上報名系統", "login.asp"))
//insDoc(foldersTree, gLnk(0, "資料下載", "資料下載"))
//insDoc(foldersTree, gLnk(0, "留言版", "Board/board.asp"))
initializeDocument()
</SCRIPT> 

&nbsp;<br> 
<p align="center"><font size="3" color="#000000">人氣指數：<b></font><font color="FE5AFB"><%=count%> </font></b><font size="3" color="#000000">人 </p>
<%
dim countStr(6)
	countPic=""
	For i = 1 to len(count)
		countStr(i) = mid(count,i,1)
		countPic=countPic+"<img src='images/digit/" & countStr(i) & ".gif' wide=16 height=20>"
	Next
	countPic="<p align='center'>瀏覽人數" & countPic & "</p>"
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
<p align="center">建議解析度：<br> 
	<font color="Red">1024 x 768</font></p>
</body> 
</html>

