<!--#InClude File="sql_conn.asp" -->
<% 
if request("logout2")="yes" then
	Session("LoginID2")=""
	Session("Password2")=""
	Session("LogonLevel")=0
	Session("UsrLevel")=0
	Session("TableType1_2")=0
	Session("OfficeType")=0
end if

sql="Select * From condition Where ConditionID=1"
set rs = conn.Execute(sql)
endday=rs("EndDate")
if len(Request.Form("cmd1"))>0 then              
    muserid=Trim(Request.Form("id"))
    mpasswd=Request.Form("password")
    sql="Select * From login Where LoginName='"+muserid+"' and Password='"+mpasswd+"'"
    set rs = conn.Execute(sql)              
	if rs.eof then              
%>
	<script language="JavaScript">
		alert("\n�п�J���T���b���B�K�X�I");
		history.go( -1 )
	</script>
<%
		flag=0              
	else              
		Session("LoginID2") = rs("LoginName")
		Session("Password2") = mpasswd
		Session("LogonLevel") = rs("RefType")
		Session("UsrLevel") = rs("usrLevel")
		LoginTimes = rs("LoginTimes")+1
		Session("TableType1_2") = Request.Form("TableType1")
		If Session("UsrLevel")=99 then
			sql1="Update Login Set LoginDate='" & date & " " & Hour(time( )) & right(time( ),6) & "'"
			sql1=sql1+",LoginIP='" & Request.ServerVariables("REMOTE_HOST") & "',LoginTimes=" & LoginTimes & " Where LoginName='"+muserid+"'"
			set rs1 = conn.Execute(sql1)
			Response.Redirect "verify2.asp"
		else
%>
	<script language="JavaScript">
		alert("\n�z���v�������A\n�ݬ��޲z�̤~�i�n�J�I");
		history.go( -1 )
	</script>
<%
			flag=0              
		end if 
	end if
else              
	flag=0              
end if              
%>
<%if flag=0 then%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>�]���޲z�̵n�J�b��</title>
</head>
<script language="Javascript"> 
function checkvalue() { 
	var tmp	= login.id.value.toLowerCase(); 
	var passwordtmp	= login.password.value.toLowerCase(); 
	if(tmp.length == 0 ){ 
	alert("�b�����i���ť�!"); 
	return false; 
	} 
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
<body onload="document.login.id.focus()">
<h2 align="center">
  <font face="�з���" size=5>�]���޲z�̵n�J�b��<br></font>
</h2>
<table border="0" width="100%">
  <tr>
    <td width="10%"></td>
    <td width="90%">
      <form method="POST" autocomplete="off" name="login" action="login2.asp">
        <table border="0" width="100%">
          <tr>
            <td width="40%">
              <p align="right">�b�� (�����Ҹ��X )�G</td>
            <td width="60%"><input type="text" name="id" size="20" maxlength="10" value="<%=session("loginid2")%>" onFocus="this.style.backgroundColor='#ECFFFF'" onBlur="this.style.backgroundColor='#FFFFFF'" onMouseOver="this.focus()"></td>
          </tr>
          <tr>
            <td width="40%">
              <p align="right">�K�X (�Ĥ@���Цۦ��J )�G</td>
            <td width="60%"><input type="password" name="password" size="20" maxlength="20" onFocus="this.style.backgroundColor='#ECFFFF'" onBlur="this.style.backgroundColor='#FFFFFF'" onMouseOver="this.focus()"></td>
          </tr>
          <tr>
            <td width="40%" align="right">��ܧ�����O�G</td>
            <td width="60%"><select name="TableType1" size="1" ID="Select1">
              <option value=1>���Ǧ~�׻P�Ǵ��O</option>
            </select></td>
          </tr>
          <tr>
            <td width="40%" height="60">
            </td>
            <td width="60%" height="60"><input type="submit" value="�T�w" name="cmd1" onclick="return checkvalue()"> 
              <input type="reset" value="���s��J" name="B2"></td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
<p align=center><font face="�з���" color=blue size=4>�����]����g���ɶ�������<%=endday%></font></p>
<p align=center><font color=red style="font-size: 10pt">�����G�K�X�Ĥ@���ۦ��J��Y�ͮġA�аO�o�z���K�X�A�H�K���n�J�ϥΡC</font></p>
<p align=center><font color=gray style="font-size: 12pt">�D��ĳ�ϥ� 1024 �� 768 �ѪR��</font></p>
<%end if%>
</body>
</html>
