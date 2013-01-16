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
		alert("\n請輸入正確的帳號、密碼！");
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
		alert("\n您的權限不足，\n需為管理者才可登入！");
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
<title>夜輔管理者登入帳號</title>
</head>
<script language="Javascript"> 
function checkvalue() { 
	var tmp	= login.id.value.toLowerCase(); 
	var passwordtmp	= login.password.value.toLowerCase(); 
	if(tmp.length == 0 ){ 
	alert("帳號不可為空白!"); 
	return false; 
	} 
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
<body onload="document.login.id.focus()">
<h2 align="center">
  <font face="標楷體" size=5>夜輔管理者登入帳號<br></font>
</h2>
<table border="0" width="100%">
  <tr>
    <td width="10%"></td>
    <td width="90%">
      <form method="POST" autocomplete="off" name="login" action="login2.asp">
        <table border="0" width="100%">
          <tr>
            <td width="40%">
              <p align="right">帳號 (身分證號碼 )：</td>
            <td width="60%"><input type="text" name="id" size="20" maxlength="10" value="<%=session("loginid2")%>" onFocus="this.style.backgroundColor='#ECFFFF'" onBlur="this.style.backgroundColor='#FFFFFF'" onMouseOver="this.focus()"></td>
          </tr>
          <tr>
            <td width="40%">
              <p align="right">密碼 (第一次請自行輸入 )：</td>
            <td width="60%"><input type="password" name="password" size="20" maxlength="20" onFocus="this.style.backgroundColor='#ECFFFF'" onBlur="this.style.backgroundColor='#FFFFFF'" onMouseOver="this.focus()"></td>
          </tr>
          <tr>
            <td width="40%" align="right">選擇更改類別：</td>
            <td width="60%"><select name="TableType1" size="1" ID="Select1">
              <option value=1>更改學年度與學期別</option>
            </select></td>
          </tr>
          <tr>
            <td width="40%" height="60">
            </td>
            <td width="60%" height="60"><input type="submit" value="確定" name="cmd1" onclick="return checkvalue()"> 
              <input type="reset" value="重新輸入" name="B2"></td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
<p align=center><font face="標楷體" color=blue size=4>本次夜輔填寫之時間期限至<%=endday%></font></p>
<p align=center><font color=red style="font-size: 10pt">附註：密碼第一次自行輸入後即生效，請記得您的密碼，以便日後登入使用。</font></p>
<p align=center><font color=gray style="font-size: 12pt">．建議使用 1024 × 768 解析度</font></p>
<%end if%>
</body>
</html>
