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
		<title>��ܿ�׽ҵ{���</title>
	</head>
	<body bgcolor="#C7EDCC">
<Br><Br>
<%
	Account=Session("LoginID")
	Password=Session("Password")
	UserLevel=Session("LogonLevel")
	' ���o �n�J�b���W��(login) �Z�Žs��(Class)
	sql="Select * from PermRec INNER JOIN Login on RefID=PermRecID INNER JOIN Class on PermRec.ClassID=Class.ClassID where SNum='" & Account & "'"
	set rs=conn.Execute(sql)
	gradeyear=rs("ClassYear")+1
	classname=rs("ClassName")
	sql1="Select * From Condition Where ConditionID=1"
	set rs1=conn.Execute(sql1)
	
    if Request.Form("nextsubmit")="�T�w" then
		if Request.Form("Subject")=0 then
			Response.Write("<td><p align='center'><span style='font-size: 18pt; color:#FF0000; font-family: �з���'><strong>�п�ܤ@�Ӭ�ث�A�A���T�w�I</strong></span><br><br></td>")
		else
			sql_Course0="Select * from CanSelectCourse where CourseID=" & Request.Form("Subject")
			set rs_Course0=conn.Execute(sql_Course0)
			'--CourseID �ҵ{�N��   
			'--SubjectValue �s�ܧ󪺽ҵ{�N�� 
			'--Selected �s�ҵ{���w��H��
			'--SelectedValue �s�ܧ󪺽ҵ{�N���w�諸�H��
			SubjectValue=rs_Course0("CourseID")
			SelectedValue=rs_Course0("Selected")
			'Response.Write("sql_Course0=" & SubjectValue & "," & SelectedValue)
	
			if (Not isnull(rs("CourseSelectID")) or rs("CourseSelectID")<>"") and rs("CourseSelectID")>0 then
				'--�½ҵ{�w��� (rs("CourseSelectID") != null)
				if rs("CourseSelectID")<>SubjectValue then
					' ��s�ǥͷs�ҵ{
					sql_update1="Update PermRec Set CourseSelectID=" & SubjectValue
					sql_update1=sql_update1+" Where SNum='" & Account & "'"
					set rs_update1=conn.Execute(sql_update1)
					'Response.Write("sql_update1=" & sql_update1)
					' --- �]���ܧ��� �G�s�ҵ{���w��ҤH�� +1 
					sql_update2="Update CanSelectCourse Set Selected=" & SelectedValue+1
					sql_update2=sql_update2+" Where CourseID=" & SubjectValue
					set rs_update2=conn.Execute(sql_update2)
					'Response.Write("sql_update2=" & sql_update2)
					
					' --- ����n�R�����e�w�諸�ҵ{�H�� -1
					sql_Course_old="Select * from CanSelectCourse where CourseID=" & rs("CourseSelectID")
					set rs_Course_old=conn.Execute(sql_Course_old)
					'SelectedValue_old �½ҵ{�ҵ{���w��ҤH��
					SelectedValue_old=rs_Course_old("Selected")
					'�]���ܧ��� �G�n�R�����e�w�諸�ҵ{���w��ҤH��-1
					sql_update3="Update CanSelectCourse Set Selected=" & SelectedValue_old-1
					sql_update3=sql_update3+" Where CourseID=" & rs("CourseSelectID")
					set rs_update3=conn.Execute(sql_update3)
					'Response.Write("sql_update3=" & sql_update3)
				end if
			else
				' --- ���e����� (CourseSelectID == null) 
				' --- �ݩ�Ĥ@�����
				sql_update1="Update PermRec Set CourseSelectID=" & SubjectValue
				sql_update1=sql_update1+" Where SNum='" & Account & "'"
				set rs_update1=conn.Execute(sql_update1)
				'Response.Write("0sql_update1=" & sql_update1)
				sql_update2="Update CanSelectCourse Set Selected=" & SelectedValue+1
				sql_update2=sql_update2+" Where CourseID=" & SubjectValue
				set rs_update2=conn.Execute(sql_update2)
				'Response.Write("0sql_update2=" & sql_update2)
			end if
			Response.Write("<td><p align='center'><span style='font-size: 18pt; color:#FF0000; font-family: �з���'><strong>�z�w������׽ҵ{���<br>�z��ܪ��ҵ{���y"+rs_Course0("CourseName")+"�z<br>�T�w�����Ы�����n�X�Y�i�I</strong></span><br><br></td>")
		end if
	end if
	sql="Select * from PermRec INNER JOIN Login on RefID=PermRecID INNER JOIN Class on PermRec.ClassID=Class.ClassID where SNum='" & Account & "'"
	set rs=conn.Execute(sql)
	
	Response.Write("<table border='0' width='100%'>")
	Response.Write("<form method='POST' name='Select' action='main.asp'>")
	Response.Write("<tr>")
	Response.Write("<td colspan=2><p align='center'><span style='font-size: 18pt; font-family: �з���'> " & rs1("SchoolYear") & "�Ǧ~�ײ�" & rs1("Semester") & "�Ǵ���" & replace(replace(replace(gradeyear,"1","�@"),"2","�G"),"3","�T") & "��׽ҵ{�@����</span></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td colspan=2><span style='font-size: 15pt; font-family: �з���'> �Z�šG" & rs("ClassName") & "�@�@�@�@�@�Ǹ��G" & Account & "�@�@�@�@�@�@�m�W�G" & rs("Name") & "</span></td>")
	Response.Write("</tr>")

	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
    Response.Write("<td colspan=2 align=center><font color=#000080>��ܽҵ{ ( ���I����ѥ[����׽ҵ{�A�M����u�T�w�v�s )</font></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	if gradeyear=2 then
		Response.Write("<td colspan=2> �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@<a href='101�~�פG�~�ſ�׽ҵ{.doc' target='_blank'>�U�� 101�~�פG�~�ſ�׽ҵ{����</td>")
	elseif gradeyear=3 then
		Response.Write("<td colspan=2> �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@<a href='101�~�פT�~�ſ�׽ҵ{.doc' target='_blank'>�U�� 101�~�פT�~�ſ�׽ҵ{����</td>")
	end if
	
	Response.Write("</tr>")
	Response.Write("<table border='1' width='100%'>")
	Response.Write("<tr>")
	Response.Write("<td align=center bgcolor=#FFCCFF width='3%'></td>")
	Response.Write("<td align=center bgcolor=#FFCCFF>���</td>")
	Response.Write("<td align=center bgcolor=#FFCCFF width='3%'>�Ǥ�</td>")
	Response.Write("<td align=center bgcolor=#FFCCFF>�ҵ{����</td>")
	Response.Write("<td align=center bgcolor=#FFCCFF width='5%'>�i��H��</td>")
	Response.Write("<td align=center bgcolor=#FFCCFF width='5%'>�w��H��</td>")
	Response.Write("<td align=center bgcolor=#FFCCFF width='5%'>�B�����p</td>")
	Response.Write("</tr>")
	
	' �̾ǥ�"��O"�P"�~��" �h �z�� �ǥͥi�H��ܪ� ��׽ҵ{
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
			flag="�B��"
		end if
		if rs("CourseSelectID")=CourseID then
			Response.Write("<input type='radio' value=" & CourseID & " name='Subject' checked></td>")
		else
			if flag="�B��" then
				Response.Write("<input type='radio' value=" & CourseID & " name='Subject' disabled></td>")
			else
				Response.Write("<input type='radio' value=" & CourseID & " name='Subject'></td>")
			end if
		end if 
		
		des_string=rs_Course("Description")
		tempArr = Split(des_string,"�C")
		
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
    Response.Write("<br><p align='center'><td width='50%' height='60'><input class='button' type='submit' value='�T�w' name='nextsubmit' onclick='return checkvalue()' ID='Submit1'>")
    Response.Write("</td>")
    Response.Write("</tr>")
    
	Response.Write("</form>")
	Response.Write("</table>")
	Response.Write("</body>")
%> 
</html>
