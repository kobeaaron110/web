<html>
	<head>
		<title>線上預選學程</title>
		<meta http-equiv="Content-Language" content="zh-tw">
		<meta http-equiv="Content-Type" content="text/html; charset=big5">
		<script language="Javascript"> 
function checkvalue() { 
	var tmp	= login.id.value.toLowerCase(); 
	if(tmp.length < 10 ){ 
	alert("請輸入10碼的帳號 (身分證號碼)!"); 
	return false; 
	} 
	if (tmp.substring(0,1) < "a" || tmp.substring(0,1) > "z" ){ 
	alert("帳號請使用英文字母開頭"); 
	return false; 
	} 
	var passwordtmp	= login.password.value.toLowerCase(); 
	if(passwordtmp.length == 0 ){ 
	alert("密碼不可為空白!"); 
	return false; 
	} 
	for(p = 0 ; p < passwordtmp.length ; p++){ 
	passwordtmp2 = passwordtmp.substring(p,p+1); 
		if((passwordtmp2 < "a" || passwordtmp2 > "z") && (passwordtmp2 < "0" || passwordtmp2 > "9")){ 
		alert("密碼請使用英文字母和數字組成，請勿使用中文字"); 
		return false; 
		} 
	} 
	return true; 
} 
		</script>
	</head>
	<body onload="document.login.id.focus()">
		<H2 align="center">
			<font face="標楷體" size="6">台中市明德女中夜輔暨晚自習線上選擇</font><FONT face="標楷體" size="5"><br>
		</H2></FONT>
		<table border="0" width="100%" ID="Table1">
			<tr>
				<td width="100%">
			<tr>
				<td width="100%" align="center" height="60"><FONT face="標楷體" size="5"><A href="http://md1.mdhs.tc.edu.tw/selectncourse/login.asp">於校內選擇課程</A></FONT></td>
			</tr>
			<tr>
				<td width="10%" align="center" height="15"></td>
			</tr>
			<tr>
				<td width="100%" align="center" height="60"><FONT face="標楷體" size="5"><A href="http://md1.mdhs.tc.edu.tw/selectncourse/login.asp">於校外選擇課程</A></FONT></td>
			</tr>
			</td>
			<TR>
				<td width="0%"></td>
			</TR>
		</table>
		<p align="center"><font color="red" style="FONT-SIZE: 10pt" size="3">附註：<FONT 
color=blue>於校內選學程，請選擇校內選擇科目</FONT>；於校外選學程，請選擇校外選擇科目。</font></p>
		<p align="center"><font color="gray" style="FONT-SIZE: 12pt">．建議使用 1024 × 768 解析度</font></p>
	</body>
</html>