<% Session.codepage=65001%>
<% Response.CharSet="utf-8" %>
<!--#InClude File="sql_conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Language" content="zh-tw">
		<meta http-equiv="Content-Type" content="text/html; charset=big5" />
		<STYLE type="text/css">
		A {text-decoration:none}
		A:hover {color: #FF0000; text-decoration: underline}
		</STYLE>
		<title>選擇選修課程科目</title>
	</head>
	<body bgcolor="#C7EDCC">
<Br><Br>
<%
	Account=Session("LoginID")
	Password=Session("Password")
	UserLevel=Session("LogonLevel")
	' 取得 登入帳號名稱(login) 班級編號(Class)
	sql="Select * from PermRec INNER JOIN Login on RefID=PermRecID INNER JOIN Class on PermRec.ClassID=Class.ClassID where SNum='" & Account & "'"
	set rs=conn.Execute(sql)
	gradeyear=rs("ClassYear")+1
	classname=rs("ClassName")
	sql1="Select * From Condition Where ConditionID=1"
	set rs1=conn.Execute(sql1)
	
    if Request.Form("nextsubmit")="確定" then
		if Request.Form("Subject")=0 then
			Response.Write("<td><p align='center'><span style='font-size: 18pt; color:#FF0000; font-family: 標楷體'><strong>請選擇一個科目後，再按確定！</strong></span><br><br></td>")
		else
			sql_Course0="Select * from CanSelectCourse where CourseID=" & Request.Form("Subject")
			set rs_Course0=conn.Execute(sql_Course0)
			'--CourseID 課程代號   
			'--SubjectValue 新變更的課程代號 
			'--Selected 新課程的已選人數
			'--SelectedValue 新變更的課程代號已選的人數
			SubjectValue=rs_Course0("CourseID")
			SelectedValue=rs_Course0("Selected")
			'Response.Write("sql_Course0=" & SubjectValue & "," & SelectedValue)
	
			if (Not isnull(rs("CourseSelectID")) or rs("CourseSelectID")<>"") and rs("CourseSelectID")>0 then
				'--舊課程已選課 (rs("CourseSelectID") != null)
				if rs("CourseSelectID")<>SubjectValue then
					' 更新學生新課程
					sql_update1="Update PermRec Set CourseSelectID=" & SubjectValue
					sql_update1=sql_update1+" Where SNum='" & Account & "'"
					set rs_update1=conn.Execute(sql_update1)
					'Response.Write("sql_update1=" & sql_update1)
					' --- 因為變更選課 故新課程的已選課人數 +1 
					sql_update2="Update CanSelectCourse Set Selected=" & SelectedValue+1
					sql_update2=sql_update2+" Where CourseID=" & SubjectValue
					set rs_update2=conn.Execute(sql_update2)
					'Response.Write("sql_update2=" & sql_update2)
					
					' --- 之後要刪掉之前已選的課程人數 -1
					sql_Course_old="Select * from CanSelectCourse where CourseID=" & rs("CourseSelectID")
					set rs_Course_old=conn.Execute(sql_Course_old)
					'SelectedValue_old 舊課程課程的已選課人數
					SelectedValue_old=rs_Course_old("Selected")
					'因為變更選課 故要刪掉之前已選的課程的已選課人數-1
					sql_update3="Update CanSelectCourse Set Selected=" & SelectedValue_old-1
					sql_update3=sql_update3+" Where CourseID=" & rs("CourseSelectID")
					set rs_update3=conn.Execute(sql_update3)
					'Response.Write("sql_update3=" & sql_update3)
				end if
			else
				' --- 之前未選課 (CourseSelectID == null) 
				' --- 屬於第一次選課
				sql_update1="Update PermRec Set CourseSelectID=" & SubjectValue
				sql_update1=sql_update1+" Where SNum='" & Account & "'"
				set rs_update1=conn.Execute(sql_update1)
				'Response.Write("0sql_update1=" & sql_update1)
				sql_update2="Update CanSelectCourse Set Selected=" & SelectedValue+1
				sql_update2=sql_update2+" Where CourseID=" & SubjectValue
				set rs_update2=conn.Execute(sql_update2)
				'Response.Write("0sql_update2=" & sql_update2)
			end if
			Response.Write("<td><p align='center'><span style='font-size: 18pt; color:#FF0000; font-family: 標楷體'><strong>您已完成選修課程選擇<br>您選擇的課程為『"+rs_Course0("CourseName")+"』<br>確定完成請按左方登出即可！</strong></span><br><br></td>")
		end if
	end if
	sql="Select * from PermRec INNER JOIN Login on RefID=PermRecID INNER JOIN Class on PermRec.ClassID=Class.ClassID where SNum='" & Account & "'"
	set rs=conn.Execute(sql)
	
	Response.Write("<table border='0' width='100%'>")
	Response.Write("<form method='POST' name='Select' action='main.asp'>")
	Response.Write("<tr>")
	Response.Write("<td colspan=2><p align='center'><span style='font-size: 18pt; font-family: 標楷體'> " & rs1("SchoolYear") & "學年度第" & rs1("Semester") & "學期高" & replace(replace(replace(gradeyear,"1","一"),"2","二"),"3","三") & "選修課程一覽表</span></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td colspan=2><span style='font-size: 15pt; font-family: 標楷體'> 班級：" & rs("ClassName") & "　　　　　學號：" & Account & "　　　　　　姓名：" & rs("Name") & "</span></td>")
	Response.Write("</tr>")

	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
    Response.Write("<td colspan=2 align=center><font color=#000080>選擇課程 ( 請點選欲參加之選修課程，然後按「確定」鈕 )</font></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	if gradeyear=2 then
		Response.Write("<td colspan=2> 　　　　　　　　　　　　　　　　　　　<a href='101年度二年級選修課程.doc' target='_blank'>下載 101年度二年級選修課程說明</td>")
	elseif gradeyear=3 then
		Response.Write("<td colspan=2> 　　　　　　　　　　　　　　　　　　　<a href='101年度三年級選修課程.doc' target='_blank'>下載 101年度三年級選修課程說明</td>")
	end if
	
	Response.Write("</tr>")
	Response.Write("<table border='1' width='100%'>")
	Response.Write("<tr>")
	Response.Write("<td align=center bgcolor=#FFCCFF width='3%'></td>")
	Response.Write("<td align=center bgcolor=#FFCCFF>科目</td>")
	Response.Write("<td align=center bgcolor=#FFCCFF width='3%'>學分</td>")
	Response.Write("<td align=center bgcolor=#FFCCFF>課程說明</td>")
	Response.Write("<td align=center bgcolor=#FFCCFF width='5%'>可選人數</td>")
	Response.Write("<td align=center bgcolor=#FFCCFF width='5%'>已選人數</td>")
	Response.Write("<td align=center bgcolor=#FFCCFF width='5%'>額滿狀況</td>")
	Response.Write("</tr>")
	
	' 依學生"科別"與"年級" 去 篩選 學生可以選擇的 選修課程
	sql_Course="Select * from CanSelectCourse where (DeptID1=" & rs("DeptID") & " or DeptID2=" & rs("DeptID") & ") and ClassYear=" & gradeyear
	set rs_Course=conn.Execute(sql_Course)
	'Response.Write("<td bgcolor=#EBEBEB>" & sql_Course & "</td>")
	
	While not rs_Course.EOF
		Response.Write("<tr>")
		i=i+1
		if (i mod 2)=1 then 
			ColorID="#EBEBEB"
		else 
			ColorID="#FBFAFA"
		end if 
		
		flag=""
		Response.Write("<td bgcolor=" & ColorID & ">")
		CourseID=rs_Course("CourseID")
		if rs_Course("Selected")>=rs_Course("StuLimit") then
			flag="額滿"
		end if
		if rs("CourseSelectID")=CourseID then
			Response.Write("<input type='radio' value=" & CourseID & " name='Subject' checked></td>")
		else
			if flag="額滿" then
				Response.Write("<input type='radio' value=" & CourseID & " name='Subject' disabled></td>")
			else
				Response.Write("<input type='radio' value=" & CourseID & " name='Subject'></td>")
			end if
		end if 
		
		des_string=rs_Course("Description")
		tempArr = Split(des_string,"。")
		
		Response.Write("<td bgcolor=" & ColorID & ">")
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'>" & rs_Course("CourseName") & "</span></font></td>")
		Response.Write("<td bgcolor=" & ColorID & ">")
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'>" & rs_Course("Credit") & "</span></font></td>")
		Response.Write("<td bgcolor=" & ColorID & "><p align='left'><font color='#000000'><span style='font-size: 12pt'>")
		'Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'>" & rs_Course("Description") & "</span></font></td>")
		if rs("DeptID")>=9 then
			For j=0 to Ubound(tempArr)-1
				Response.Write(tempArr(j) & "<br>")
			Next
		else
			Response.Write(des_string)
		end if
		Response.Write("</td>")
		Response.Write("<td bgcolor=" & ColorID & ">")
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'>" & rs_Course("StuLimit") & "</span></font></td>")
		Response.Write("<td bgcolor=" & ColorID & ">")
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'>" & rs_Course("Selected") & "</span></font></td>")
		Response.Write("<td bgcolor=" & ColorID & ">")
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'>" & flag & "</span></font></td>")
		Response.Write("</tr>")
		rs_Course.MoveNext
	Wend
						
	
	Response.Write("</table>")
	
	
	
    Response.Write("<tr>")
    Response.Write("<br><p align='center'><td width='50%' height='60'><input class='button' type='submit' value='確定' name='nextsubmit' onclick='return checkvalue()' ID='Submit1'>")
    Response.Write("</td>")
    Response.Write("</tr>")
    
	Response.Write("</form>")
	Response.Write("</table>")
	Response.Write("</body>")
%> 
</html>
