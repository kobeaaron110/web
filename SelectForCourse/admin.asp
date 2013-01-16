<!--#InClude File="sql_conn.asp" -->
<%
sql="Select * From Condition Where ConditionID=1"
set rs = conn.Execute(sql)
if Not rs.eof then              
    if mid(left(rs("EndDate"),10),7,1)="/" then
		if len(left(rs("EndDate"),10))=8 then
			SysEndDate=mid(left(rs("EndDate"),10),1,4) & "0" & mid(left(rs("EndDate"),10),6,1) & "0" & mid(left(rs("EndDate"),10),8,1)
		else
			SysEndDate=mid(left(rs("EndDate"),10),1,4) & "0" & mid(left(rs("EndDate"),10),6,1) & mid(left(rs("EndDate"),10),8,2)
		end if
    end if
    if mid(left(rs("EndDate"),10),8,1)="/" then
		if len(left(rs("EndDate"),10))=9 then
			SysEndDate=mid(left(rs("EndDate"),10),1,4) & mid(left(rs("EndDate"),10),6,2) & "0" & mid(left(rs("EndDate"),10),9,1)
		else
			SysEndDate=mid(left(rs("EndDate"),10),1,4) & mid(left(rs("EndDate"),10),6,2) & mid(left(rs("EndDate"),10),9,2)
		end if
    end if
end if

if len(Request.Form("B1"))>0 then              
    SchoolYear=Trim(Request.Form("SY"))
    Semester=Request.Form("SM")
    EndDate=Request.Form("EndDay1")
    if rs.eof then              
		sql1="Insert Into Condition (SchoolYear,Semester,IP,Date) values (" & SchoolYear & "," & Semester & ",'" & Request.ServerVariables("HTTP_X_FORWARDED_FOR") & "','" & Date & "')"
		set rs1 = conn.Execute(sql1)
	else
	  if rs("SchoolYear")=int(SchoolYear) And rs("Semester")=int(Semester) then
		if SysEndDate=EndDate then
%>
	<script language="JavaScript">
		alert("\n您選擇的學年度、學期別與填寫截止日期\n與目前設定相同，無須修改！\n");
	</script>
<%
		else
			sql1="Update Condition Set EndDate='" & EndDate & "',IP='" & Request.ServerVariables("REMOTE_HOST") & "',Date='" & Date & "' Where ConditionID=1"
			set rs1 = conn.Execute(sql1)
			sql="Select * From Condition Where ConditionID=1"
			set rs = conn.Execute(sql)
			if mid(left(rs("EndDate"),10),7,1)="/" then
				if len(left(rs("EndDate"),10))=8 then
					SysEndDate=mid(left(rs("EndDate"),10),1,4) & "0" & mid(left(rs("EndDate"),10),6,1) & "0" & mid(left(rs("EndDate"),10),8,1)
				else
					SysEndDate=mid(left(rs("EndDate"),10),1,4) & "0" & mid(left(rs("EndDate"),10),6,1) & mid(left(rs("EndDate"),10),8,2)
				end if
			end if
			if mid(left(rs("EndDate"),10),8,1)="/" then
				if len(left(rs("EndDate"),10))=9 then
					SysEndDate=mid(left(rs("EndDate"),10),1,4) & mid(left(rs("EndDate"),10),6,2) & "0" & mid(left(rs("EndDate"),10),9,1)
				else
					SysEndDate=mid(left(rs("EndDate"),10),1,4) & mid(left(rs("EndDate"),10),6,2) & mid(left(rs("EndDate"),10),9,2)
				end if
			end if
%>
	<script language="JavaScript">
		alert("\n已更新填寫截止日期完成！\n");
	</script>
<%
		end if
	  else
		sql1="Select * From Condition Where ConditionID=2"
		set rs1 = conn.Execute(sql1)
		'response.write("<center><font color=#FF0000>"+rs("IP")+Request.ServerVariables("HTTP_VIA")+"</font></center><p>")              
		if rs1.eof then
			sql1="Insert Into Condition (SchoolYear,Semester,EndDate,IP,Date) values (" & rs("SchoolYear") & "," & rs("Semester")  & ",'" & rs("EndDate")& "','" & rs("IP") & "','" & rs("Date") & "')"
			set rs1 = conn.Execute(sql1)
		else
			sql1="Update Condition Set SchoolYear=" & rs("SchoolYear") & ",Semester=" & rs("Semester") & ",EndDate='" & rs("EndDate") & "',IP='" & rs("IP") & "',Date='" & rs("Date") & "' Where ConditionID=2"
			set rs1 = conn.Execute(sql1)
		end if
		sql1="Update Condition Set SchoolYear=" & SchoolYear & ",Semester=" & Semester & ",EndDate='" & EndDate & "',IP='" & Request.ServerVariables("REMOTE_HOST") & "',Date='" & Date & "' Where ConditionID=1"
		set rs1 = conn.Execute(sql1)
%>
	<script language="JavaScript">
		alert("\n已更新完成！\n");
	</script>
<%
		sql="Select * From Condition Where ConditionID=1"
		set rs = conn.Execute(sql)
		if mid(left(rs("EndDate"),10),7,1)="/" then
			if len(left(rs("EndDate"),10))=8 then
				SysEndDate=mid(left(rs("EndDate"),10),1,4) & "0" & mid(left(rs("EndDate"),10),6,1) & "0" & mid(left(rs("EndDate"),10),8,1)
			else
				SysEndDate=mid(left(rs("EndDate"),10),1,4) & "0" & mid(left(rs("EndDate"),10),6,1) & mid(left(rs("EndDate"),10),8,2)
			end if
		end if
		if mid(left(rs("EndDate"),10),8,1)="/" then
			if len(left(rs("EndDate"),10))=9 then
				SysEndDate=mid(left(rs("EndDate"),10),1,4) & mid(left(rs("EndDate"),10),6,2) & "0" & mid(left(rs("EndDate"),10),9,1)
			else
				SysEndDate=mid(left(rs("EndDate"),10),1,4) & mid(left(rs("EndDate"),10),6,2) & mid(left(rs("EndDate"),10),9,2)
			end if
		end if
	  end if             
	end if
else              
	flag=0              
end if              
%>
<script language="Javascript"> 
function checkvalue() { 
	RightNow = new Date();
	var tmp	= admins.EndDay1.value.toLowerCase(); 
	if(tmp.length == 0 ){ 
	alert("有效日期不可為空白!"); 
	return false; 
	} 
	if(tmp.length < 8 ){ 
	alert("請輸入8碼的有效日期!"); 
	return false; 
	} 
	for(p = 0 ; p < tmp.length ; p++){ 
	tmp2 = tmp.substring(p,p+1); 
		if (tmp2 < "0" || tmp2 > "9"){ 
		alert("有效日期請使用數字，請勿使用文字"); 
		return false; 
		} 
	} 
	tmpYY = tmp.substring(0,4); 
		if (tmpYY != RightNow.getFullYear()){ 
		alert("有效日期年份需為"+RightNow.getFullYear()+"之數字!"); 
		return false; 
		} 
	tmpMM = tmp.substring(4,6); 
	if(tmpMM < 0 || tmpMM > 12 ){ 
		alert("有效日期月份需為1至12之數字!"); 
		return false; 
	} 
	tmpDD = tmp.substring(6,8); 
	if(tmpMM == 1 || tmpMM == 3 || tmpMM == 5 || tmpMM == 7 || tmpMM == 8 || tmpMM == 10 || tmpMM == 12 ){
		if(tmpDD < 0 || tmpDD > 31 ){
		alert("有效日期日份需為1至31之數字!");
		return false;
		}}
	else if(tmpMM == 2 ){
		if(tmpDD < 0 || tmpDD > 29 ){
		alert("有效日期日份需為1至29之數字!");
		return false;
		}}
	else {
		if(tmpDD < 0 || tmpDD > 30 ){
		alert("有效日期日份需為1至30之數字!");
		return false;
		}
	}
	return true; 
} 
</script>  
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<STYLE type="text/css">
A {text-decoration:none}
A:hover {color: #FF0000; text-decoration: underline}
</STYLE>
<title>管理者管理資料</title>
</head>
<body link="#0000FF" vlink="#0000FF" alink="#FF0000" style="font-size: 12pt">
<p align="center"><font color="#0000FF"><b>夜輔時間變更選單</b></font></p>
<form method="POST" name="admins" action="admin.asp">
  <div align="center">
    <center>
    <table border="0">
      <tr>
        <td width="50%" align="right">選擇夜輔之學年度：</td>
        <td width="50%"><select name="SY" size="1">
<% for i=96 to 120 
      if rs("SchoolYear")=i then %>
          <option value=<%=i%> selected><%=i%></option>
<%    else %>
          <option value=<%=i%>><%=i%></option>
<%    end if
   next %>
          </select></td>
      </tr>
      <tr>
        <td width="50%" align="right">選擇夜輔之學期別：</td>
        <td width="50%"><select name="SM" size="1">
<% for i=1 to 2 
      if rs("Semester")=i then %>
              <option value=<%=i%> selected><%if i=1 then response.write "上學期" else response.write "下學期" end if%></option>
<%    else %>
              <option value=<%=i%>><%if i=1 then response.write "上學期" else response.write "下學期" end if%></option>
<%    end if
   next %>
          </select></td>
      </tr>
      <tr>
        <td width="50%" align="right">填寫截止日期(8位數之日期，如 20080101)：</td>
        <td width="50%"><input type="text" name="EndDay1" size="10" maxlength="8" value=<%=SysEndDate%>></td>
     </tr>
      <tr>
      </tr>
      <tr>
          <td width="45%" align="right" height="60"></td>
          <td width="50%" height="60"><input type="submit" value="確定" name="B1" onclick="return checkvalue()"> 
          </td>
      </tr>
    </table>
    </center>
  </div>
</form>
<div align="center">
  <center>
  <table border="0" style="font-size: 10pt">
    <tr>
      <td><a href="login2.asp?logout2=yes">結束更新作業時請按這裡登出</a></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>
