<% Session.CodePage=950 %>
<% Response.CharSet="BIG5" %>
<!--#InClude File="sql_conn.asp" -->
<%
if request("logout")="yes" then
	Session("loginid")=""
	Session("Password")=""
	Session("LogonLevel")=0
	Session("TableLevel")=0
	Session("botanizedID")=0
end if

sql="Select * From condition Where ConditionID=1"
set rs = conn.Execute(sql)
startday=rs("StartDate")
endday=rs("EndDate")

if len(Request.Form("cmd1"))>0 then              
    muserid=Trim(Request.Form("id"))
    mpasswd=Request.Form("password")
	'//-- aaron Not
	if endday<now or startday>now then 'and muserid<>"P101825460"
%>
	<script language="JavaScript">
		alert("\n不在時間期限內，\n您已無權限進入！");
		history.go( -1 )
	</script>
<%
    else
      sql1="Select * From login Where LoginName='"+muserid+"'"
      set rs1 = conn.Execute(sql1)
      if rs1.eof then              
%>
	<script language="JavaScript">
		alert("\n請輸入正確的帳號！");
		history.go( -1 )
	</script>
<%
		flag=0              
	  else
			temppasswd=rs1("Password")
		'if temppasswd="first" then
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
				sql_class="Select * from PermRec INNER JOIN Class on PermRec.ClassID=Class.ClassID where SNum='" & muserid & "'"
				set rs_class=conn.Execute(sql_class)
			    
				if rs("Active")=0 then
%>
	<script language="JavaScript">
		alert("\n您已無權限進入！");
		history.go( -1 )
	</script>
<%
					flag=0              
			    elseif left(rs_class("ClassName"),1)="普" then 'rs_class("ClassName")="普三甲" or rs_class("ClassName")="普三乙" then
%>
	<script language="JavaScript">
		alert("\n您無須填寫！");
		history.go( -1 )
	</script>
<%
					flag=0              
			    else
					Session("LoginID") = rs("LoginName")
					Session("Password") = mpasswd
					Session("LogonLevel") = rs("RefType")
					
					LoginTimes = rs("LoginTimes")+1
					sql1="Update Login Set LoginDate='" & date & " " & Hour(time( )) & right(time( ),6) & "'"
					sql1=sql1+",LoginIP='" & Request.ServerVariables("REMOTE_HOST") & "',LoginTimes=" & LoginTimes & " Where LoginName='"+muserid+"'"
					set rs1 = conn.Execute(sql1)
					Response.Redirect "verify.asp"
				end if
			end if
		'end if
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
<title>夜輔暨晚自習選擇登入帳號</title>
</head>
<script language="Javascript"> 
function checkvalue() {
	var tmp	= login.id.value.toLowerCase(); 
	if(tmp.length < 6 ){
	alert("請輸入6碼的帳號 (學號)!"); 
	return false;
	} 
	for(p = 0 ; p < 2 ; p++){ 
	tmp2 = tmp.substring(p,p+1); 
		if(tmp2 < "0" || tmp2 > "9"){ 
		alert("學號前2碼請使用數字，請勿使用其他文字"); 
		return false; 
		} 
	} 
	var passwordtmp	= login.password.value.toLowerCase(); 
	if(passwordtmp.length < 10 ){ 
	alert("請輸入10碼的密碼 (身分證號碼)!"); 
	return false; 
	} 
	if (passwordtmp.substring(0,1) < "a" || passwordtmp.substring(0,1) > "z" ){ 
	alert("密碼請使用英文字母開頭"); 
	return false; 
	} 
	for(p = 2 ; p < passwordtmp.length ; p++){ 
	tmp2 = passwordtmp.substring(p,p+1); 
		if(tmp2 < "0" || tmp2 > "9"){ 
		alert("密碼後8碼請使用數字，請勿使用其他文字"); 
		return false; 
		} 
	} 
	return true; 
}

function numcheck(id, time) {
	var re = /^[0-9]+$/;
    if (!re.test(time.value)){
		document.getElementById(id).value = time.value.substring(0, time.value.length - 1);
	}
	if (!re.test(document.getElementById(id).value)){
		document.getElementById(id).value = "";
	}
}
 
</script>  
<body onload="document.login.id.focus()">
<h2 align="center">
  <font face="標楷體" size=5>台中市明德女中夜輔暨晚自習選擇登入帳號<br></font>
	<%if date<#2012/09/05# then%>
		<font face="標楷體" color=blue size=4>本次填寫之時間期限自 <%=startday%> 至 <%=endday%><br></font>
	<%else%>
		<font face="標楷體" color=blue size=4>本次列印之時間期限 至 <%=endday%><br></font>
	<%end if%>
</h2>
<table border="0" width="100%">
  <tr>
    <td width="100%">
      <form method="POST" autocomplete="off" name="login" action="login.asp">
        <table border="0" width="100%">
          <tr>
            <td width="45%" align="right">帳號 (6位學號 )：</td>
            <td width="50%"><input type="text" id="acount"  name="id" size="10" maxlength="6" value="<%=session("loginid")%>" onFocus="this.style.backgroundColor='#ECFFFF'" onBlur="this.style.backgroundColor='#FFFFFF'" onMouseOver="this.focus()" onkeyup="numcheck(this.id,this)" ></td>
          </tr>
          <tr>
            <td width="45%" align="right">密碼 (身分證號碼，第一碼英文大小寫皆可 )：</td>
            <td width="50%"><input type="password" name="password" size="20" maxlength="10" value="<%=session("password")%>" onFocus="this.style.backgroundColor='#ECFFFF'" onBlur="this.style.backgroundColor='#FFFFFF'" onMouseOver="this.focus()"></td>
          </tr>
          <tr>
            <td width="45%" align="right" height="60"></td>
            <td width="50%" height="60"><input type="submit" value="確定" name="cmd1" onclick="checkvalue()"> 
              <input type="reset" value="重新輸入" name="B2"></td>
          </tr>
        </table>
      </form> 
    </td> 
    <td width="0%"></td>
  </tr> 
</table> 
<p align=center><font color=red style="font-size: 10pt"></font></p>
<p align=center><font color=gray style="font-size: 12pt">．建議使用 1024 × 768 解析度</font></p>
<%end if%> 
</body> 
</html> 
