<html>
	<head>
		<title>�u�W�w��ǵ{</title>
		<meta http-equiv="Content-Language" content="zh-tw">
		<meta http-equiv="Content-Type" content="text/html; charset=big5">
		<script language="Javascript"> 
function checkvalue() { 
	var tmp	= login.id.value.toLowerCase(); 
	if(tmp.length < 10 ){ 
	alert("�п�J10�X���b�� (�����Ҹ��X)!"); 
	return false; 
	} 
	if (tmp.substring(0,1) < "a" || tmp.substring(0,1) > "z" ){ 
	alert("�b���Шϥέ^��r���}�Y"); 
	return false; 
	} 
	var passwordtmp	= login.password.value.toLowerCase(); 
	if(passwordtmp.length == 0 ){ 
	alert("�K�X���i���ť�!"); 
	return false; 
	} 
	for(p = 0 ; p < passwordtmp.length ; p++){ 
	passwordtmp2 = passwordtmp.substring(p,p+1); 
		if((passwordtmp2 < "a" || passwordtmp2 > "z") && (passwordtmp2 < "0" || passwordtmp2 > "9")){ 
		alert("�K�X�Шϥέ^��r���M�Ʀr�զ��A�ФŨϥΤ���r"); 
		return false; 
		} 
	} 
	return true; 
} 
		</script>
	</head>
	<body onload="document.login.id.focus()">
		<H2 align="center">
			<font face="�з���" size="6">�x�������w�k���]���[�ߦ۲߽u�W���</font><FONT face="�з���" size="5"><br>
		</H2></FONT>
		<table border="0" width="100%" ID="Table1">
			<tr>
				<td width="100%">
			<tr>
				<td width="100%" align="center" height="60"><FONT face="�з���" size="5"><A href="http://md1.mdhs.tc.edu.tw/selectncourse/login.asp">��դ���ܽҵ{</A></FONT></td>
			</tr>
			<tr>
				<td width="10%" align="center" height="15"></td>
			</tr>
			<tr>
				<td width="100%" align="center" height="60"><FONT face="�з���" size="5"><A href="http://md1.mdhs.tc.edu.tw/selectncourse/login.asp">��ե~��ܽҵ{</A></FONT></td>
			</tr>
			</td>
			<TR>
				<td width="0%"></td>
			</TR>
		</table>
		<p align="center"><font color="red" style="FONT-SIZE: 10pt" size="3">�����G<FONT 
color=blue>��դ���ǵ{�A�п�ܮդ���ܬ��</FONT>�F��ե~��ǵ{�A�п�ܮե~��ܬ�ءC</font></p>
		<p align="center"><font color="gray" style="FONT-SIZE: 12pt">�D��ĳ�ϥ� 1024 �� 768 �ѪR��</font></p>
	</body>
</html>