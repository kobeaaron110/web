<% Session.codepage=65001%>
<% Response.CharSet="utf-8" %>
<!--#InClude File="sql_conn.asp" -->
<!--#InClude File="bus_checkvalue.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Language" content="zh-tw">
		<meta http-equiv="Content-Type" content="text/html; charset=big5" />
		<script type="text/javascript" src="jquery.min.js"></script>
		<STYLE type="text/css">
		A {text-decoration:none}
		A:hover {color: #FF0000; text-decoration: underline}
		
		.even {
			background-Color:#EBEBEB;
		} 
		.odd {
			background-Color:#FBFAFA;
		} 
		.eq {
			background-Color:#BDBDBD;
		}
		</STYLE>
		<title>��ܩ]���[�ߦ۲߰ѥ[���</title>
		
	</head>
	<body bgcolor="#C7EDCC">
<Br><Br>
<%
	'\\203.68.204.6\e$\�{��\101�~\2��\1010203\SelectNcourse\main.asp
	' rs("ProgramID1")=999 �N��@�w�n��
	' identity="pro" �N��O��O�Z
	' ---- ���W�B����
	'from CanSelectNCourse where CourseID=1"'���@��¾�ӷ~���ƾ� B�Z	  Course2              100�W
	'from CanSelectNCourse where CourseID=2"'���@��¾�a�����ƾ� A�Z	Course2 ���e ���O  80�W
	'from CanSelectNCourse where CourseID=3"'���@��¾�^��[�j�Z	 Course3  			   120�W
	
	Account=Session("LoginID")
	Password=Session("Password")
	UserLevel=Session("LogonLevel")
	sql="Select * from PermRec INNER JOIN Login on RefID=PermRecID INNER JOIN Class on PermRec.ClassID=Class.ClassID where SNum='" & Account & "' And PermRec.Active=1"
	set rs=conn.Execute(sql)
	gradeyear=rs("ClassYear")
	classname=rs("ClassName")
	course1=rs("Course1")
	course2=rs("Course2")
	course3=rs("Course3")
	course4=rs("Course4")
	course5=rs("Course5")
	course6=rs("Course6")
	course7=rs("Course7")
	' busval : BusCode
	busval=rs("Bus")
	busid=rs("BusStopID")
	if isnull(rs("Password1")) then
		identity=""
	else
		identity=rs("Password1")
	end	if
	'sql1="Select * from SystemVariable Where SysVarName = 'SelectCourseSchoolYear'"
	'set rs1=conn.Execute(sql1)
	'sql2="Select * from SystemVariable Where SysVarName = 'SelectCourseSemester'"
	'set rs2=conn.Execute(sql2)
	sql1="Select * From Condition Where ConditionID=1"
	set rs1=conn.Execute(sql1)
	
	if Request.Form("Subject1")="on" then
		Subject1=1
	else
		Subject1=0
	end if
	if Request.Form("Subject2")="on" or rs("ProgramID1")=999 or rs("ProgramID2")=999 then
		Subject2=1
	else
		Subject2=0
	end if
	if Request.Form("Subject3")="on" or rs("ProgramID3")=999 then
		Subject3=1
	else
		Subject3=0
	end if
	if Request.Form("Subject4")="on" then
		Subject4=1
	else
		Subject4=0
	end if
	if Request.Form("Subject5")="on" then
		Subject5=1
	else
		Subject5=0
	end if
	if Request.Form("Subject6")="on" then
		Subject6=1
	else
		Subject6=0
	end if
	if Request.Form("Subject7")="on" then
		Subject7=1
	else
		Subject7=0
	end if
	if busval>0 then
		sql_bus="Select * from ���W��� inner join ���ɽҮը� on BusNum=���� Where BusCode=" & busval & " order by ����"
		set rs_bus=conn.Execute(sql_bus)
	end if
	Bus=Request.Form("Bus")
	BusStopID=Request.Form("Select1")
	tel=Request.Form("Telephone")
	mobile=Request.Form("Mobile")
    if Request.Form("nextsubmit")="�U�@�B" then
		application.lock
		'if BusStopID<>"�п�ܤW�U�����W" or Bus=0 or Bus=16 then
		if BusStopID<>"�п�ܤW�U�����W" or Bus=0 then
			' ���W�B������
			' Credit: �N���ܪ�����
			sql_Course03="Select * from CanSelectNCourse where CourseID=3"
			set rs_Course03=conn.Execute(sql_Course03)
			SortValue3=rs_Course03("Credit")
			SelectedValue3=rs_Course03("Selected")
			sql_Course02="Select * from CanSelectNCourse where CourseID=2"
			set rs_Course02=conn.Execute(sql_Course02)
			SortValue2=rs_Course02("Credit")
			SelectedValue2=rs_Course02("Selected")
			sql_Course01="Select * from CanSelectNCourse where CourseID=1"
			set rs_Course01=conn.Execute(sql_Course01)
			SortValue1=rs_Course01("Credit")
			SelectedValue1=rs_Course01("Selected")
			' ---- ���W�B����
			'from CanSelectNCourse where CourseID=1"'���@��¾�ӷ~���ƾ� B�Z	  Course2
			'from CanSelectNCourse where CourseID=2"'���@��¾�a�����ƾ� A�Z	Course2
			'from CanSelectNCourse where CourseID=3"'���@��¾�^��[�j�Z	 Course3
			' -------------------------- Course2 OR Course3 ���� �B ���@�~�� ���O���� �O¾��
			' continue = 1 �N�� �n�� Update 
			continue=1
			if (Subject2=1 or Subject3=1) and gradeyear=1 and left(classname,1)<>"��" then
				continue=0
				update1=0
				update2=0
				update3=0
				if Subject3=1 then
					if rs_Course03("Selected")>=rs_Course03("StuLimit") then
						' �w�B��
						Response.Write("<td><p align='center'><span style='font-size: 18pt; color:#FF0000; font-family: �з���'><strong>�y�^��[�j�Z�z�w�B���A�Э��s��ܫ�A�~��I</strong></span><br><br></td>")						
					else
						continue=1
						update3=1
					end if
				end if
				if Subject2=1 then
					if left(classname,1)="��" or left(classname,1)="��" then
						if rs_Course02("Selected")>=rs_Course02("StuLimit") then
							Response.Write("<td><p align='center'><span style='font-size: 18pt; color:#FF0000; font-family: �з���'><strong>�y�a�����ƾ�A�Z�z�w�B���A�Э��s��ܫ�A�~��I</strong></span><br><br></td>")							
						else
							continue=1
							update2=1
						end if
					else
						if rs_Course01("Selected")>=rs_Course01("StuLimit") then
							Response.Write("<td><p align='center'><span style='font-size: 18pt; color:#FF0000; font-family: �з���'><strong>�y�ӷ~���ƾ�B�Z�z�w�B���A�Э��s��ܫ�A�~��I</strong></span><br><br></td>")					
						else
							continue=1
							update1=1
						end if
					end if
				end if
				if Subject1=1 or Subject4=1 or Subject5=1 or Subject6=1 then
					continue=1
				end if
			else
				continue=1
			end if
			' ------------------------------------------------- ���W���
			if Bus>0 then
				sql_bus="Select * from ���W��� Where �N�X=" & BusStopID
				set rs_bus=conn.Execute(sql_bus)
				BusStopName=rs_bus("���W")
			else
				BusStopID=""
			end if
			
			if tel="" or mobile="" then
				Response.Write("<td><span style='font-size: 18pt; color:#FF0000; font-family: �з���'><strong>�q�ܥ���g�A�Э��s��g�q�q�ܡA���s��ܬ�ث�F</strong></span><br><br></td>")
				Response.Write("<td><span style='font-size: 18pt; color:#FF0000; font-family: �з���'><strong>�I�ܮը�������A�A��ܤW�U�����W�A�A���U�@�B�~��I</strong></span><br><br></td>")
			else
				' -------------------------------------------------- Update CanSelectNCourse
				if continue=1 then
					'-------- update1
					if update1=1 then
						if rs("ProgramID1")=0 then
							sql_updateCourse1="Update CanSelectNCourse Set Selected=" & SelectedValue1+1 & ",Credit=" & SortValue1+1 & " Where CourseID=1"
						end if
					else
						if rs("ProgramID1")>0 and rs("ProgramID1")<999 then
							sql_updateCourse1="Update CanSelectNCourse Set Selected=" & SelectedValue1-1 & " Where CourseID=1"
						end if
					end if
					if sql_updateCourse1<>"" then
						set rs_updateCourse1=conn.Execute(sql_updateCourse1)
					end if
					'-------- update2
					if update2=1 then
						if rs("ProgramID2")=0 then
							sql_updateCourse2="Update CanSelectNCourse Set Selected=" & SelectedValue2+1 & ",Credit=" & SortValue2+1 & " Where CourseID=2"
						end if
					else
						if rs("ProgramID2")>0 and rs("ProgramID2")<999 then
							sql_updateCourse2="Update CanSelectNCourse Set Selected=" & SelectedValue2-1 & " Where CourseID=2"
						end if
					end if
					if sql_updateCourse2<>"" then
						set rs_updateCourse2=conn.Execute(sql_updateCourse2)
					end if
					'-------- update3
					if update3=1 then
						if rs("ProgramID3")=0 then
							sql_updateCourse3="Update CanSelectNCourse Set Selected=" & SelectedValue3+1 & ",Credit=" & SortValue3+1 & " Where CourseID=3"
						end if
					else
						if rs("ProgramID3")>0 and rs("ProgramID3")<999 then
							sql_updateCourse3="Update CanSelectNCourse Set Selected=" & SelectedValue3-1 & " Where CourseID=3"
						end if
					end if
					if sql_updateCourse3<>"" then
						set rs_updateCourse3=conn.Execute(sql_updateCourse3)
					end if
					
					' ---------------------------------- Update PermRec
					sql_update="Update PermRec Set Course1=" & Subject1 & ", Course2=" & Subject2 & ", Course3=" & Subject3 & ", Course4=" & Subject4
					sql_update=sql_update+", Course5=" & Subject5 & ", Course6=" & Subject6 & ", Course7=" & Subject7 & ", Bus=" & Bus & ", BusStopID='" & BusStopID
					sql_update=sql_update+"', Tel='" & tel & "', Mobile='" & mobile & "'"
					'-------- update1
					if update1=1 then
						if rs("ProgramID1")=0 then
							sql_update=sql_update+", ProgramID1=" & SortValue1+1 & ", ProgramDate1='" & now & "'"
						end if
					else
						if rs("ProgramID1")>0 and rs("ProgramID1")<999 then
							sql_update=sql_update+", ProgramID1=0, ProgramDate1='" & now & "'"
						end if
					end if
					'-------- update2
					if update2=1 then
						if rs("ProgramID2")=0 then
							sql_update=sql_update+", ProgramID2=" & SortValue2+1 & ", ProgramDate2='" & now & "'"
						end if
					else
						if rs("ProgramID2")>0 and rs("ProgramID2")<999 then
							sql_update=sql_update+", ProgramID2=0, ProgramDate2='" & now & "'"
						end if
					end if
					'-------- update3
					if update3=1 then
						if rs("ProgramID3")=0 then
							sql_update=sql_update+", ProgramID3=" & SortValue3+1 & ", ProgramDate3='" & now & "'"
						end if
					else
						if rs("ProgramID3")>0 and rs("ProgramID3")<999 then
							sql_update=sql_update+", ProgramID3=0, ProgramDate3='" & now & "'"
						end if
					end if
					sql_update=sql_update+" Where SNum='" & Account & "' And PermRec.Active=1"
					'Response.Write(BusStopID & "," & BusStopName)
					'Response.Write("sql_update=" & sql_update)
					set rs_update=conn.Execute(sql_update)
					
					application.unlock
					Response.Redirect "print.asp"
				end if
			end if
		else
			Response.Write("<td><span style='font-size: 18pt; color:#FF0000; font-family: �з���'><strong>�W�U�����W������A�Э��s��g�q�q�ܡA���s��ܬ�ث�F</strong></span><br><br></td>")
			Response.Write("<td><span style='font-size: 18pt; color:#FF0000; font-family: �з���'><strong>�I�ܮը�������A�A��ܤW�U�����W�A�A���U�@�B�~��I</strong></span><br><br></td>")
		end if
		application.unlock
	end if '�U�@�B
	
	' ------------------------------------------------------------------------------- 
	Response.Write("<table border='0' width='100%' ID='Table1'>")
	Response.Write("<form method='POST' name='Select' action='main.asp' ID='Form1'>")
	Response.Write("<tr>")
	Response.Write("<td colspan=2><p align='center'><span style='font-size: 18pt; font-family: �з���'> " & rs1("SchoolYear") & "�Ǧ~�ײ�" & rs1("Semester") & "�Ǵ��]�����ɺ[�ߦ۲߬��ʰѥ[���</span></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td colspan=2><span style='font-size: 15pt; font-family: �з���'> �Z�šG" & rs("ClassName") & "�@�@�@�@�@�Ǹ��G" & Account & "�@�@�@�@�@�@�m�W�G" & rs("Name") & "</span></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
    Response.Write("<td bgcolor='#008080' width='25%'><p align='center'><font color='#FFFFFF'><span style='font-size: 12pt'>�ж�g�p���q��</span></font></td>")
	Response.Write("<td><font color='#008080'><span style='font-size: 12pt'><INPUT type='text' value='" & rs("tel") & "' name='Telephone" & cnt & "' maxlength='20'></span></font></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
    Response.Write("<td bgcolor='#008080'><p align='center'><font color='#FFFFFF'><span style='font-size: 12pt'>�ж�g��ʹq��</span></font></td>")
	Response.Write("<td><font color='#008080'><span style='font-size: 12pt'><INPUT type='text' value='" & rs("mobile") & "' name='Mobile" & cnt & "' maxlength='10'></span></font></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	
	flag1=""
	flag2=""
	flag3=""
	sql_Course1="Select * from CanSelectNCourse where CourseID=1"'���@��¾�ӷ~���ƾ� B�Z
	set rs_Course1=conn.Execute(sql_Course1)
	sql_Course2="Select * from CanSelectNCourse where CourseID=2"'���@��¾�a�����ƾ� A�Z
	set rs_Course2=conn.Execute(sql_Course2)
	sql_Course3="Select * from CanSelectNCourse where CourseID=3"'���@��¾�^��[�j�Z
	set rs_Course3=conn.Execute(sql_Course3)
	if rs_Course1("Selected")>=rs_Course1("StuLimit") then
		flag1="�B��"
	end if
	if rs_Course2("Selected")>=rs_Course2("StuLimit") then
		flag2="�B��"
	end if
	if rs_Course3("Selected")>=rs_Course3("StuLimit") then
		flag3="�B��"
	end if
	
	if identity="pro" then 
	'	Response.Write("<td colspan=2> �ѥ[��ءG�@�@�@�@�@�@�@�@�@�@�@�@�@�@<a href='9901�^�ƽҷ~�[�]�����ɹ�I�n�I(������).doc' target='_blank'>�U�� �]�����ɺ[�ߦ۲߹�I�n�I</td>")
	elseif (gradeyear=1 or gradeyear=2) and left(classname,1)="��" then 
		Response.Write("<td colspan=2> �ѥ[��ءG�@�@�@�@�@�@�@�@�@�@�@�@�@�@<a href='10102�]�����ɹ�I�n�I(����@�G��).doc' target='_blank'>�U�� �]�����ɺ[�ߦ۲߹�I�n�I</td>")
	elseif gradeyear=1 or gradeyear=2 then 
		Response.Write("<td colspan=2> �ѥ[��ءG�@�@�@�@�@�@�@�@�@�@�@�@�@�@<a href='10102�]�����ɹ�I�n�I(���@�G¾��).doc' target='_blank'>�U�� �]�����ɺ[�ߦ۲߹�I�n�I</td>")
	elseif gradeyear=3 then 
		Response.Write("<td colspan=2> �ѥ[��ءG�@�@�@�@�@�@�@�@�@�@�@�@�@�@<a href='10102�]�����ɹ�I�n�I(���T¾��).doc' target='_blank'>�U�� �]�����ɺ[�ߦ۲߹�I�n�I</td>")
	end if
	
	' --------------------------- ��� ---------------------------
	Response.Write("<table border='1' width='100%' ID='Table2'>")
	Response.Write("<tr>")
	Response.Write("<td  width='50%'> ")
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 14pt'>�}�үZ�O�]�Цۦ�Ŀ�^</span></font></td>")
	Response.Write("<td  width='50%'> ")
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 14pt'>�A�X�ѥ[��H</span></font>")
	Response.Write("</td>")
	Response.Write("</tr>")
	
	'---------------------------- course1  �۲� ���@ ���G ���T
	' --------------------------- course1  �۲� ���@
	if course1=true or course1=false then
		Response.Write("<tr>")
		if gradeyear=1 then
			if identity="pro" then				
				Response.Write("<td><input id='Subject1' type='checkbox' name='Subject1' > ��O�ߦ۲߯Z(�P���@ �� �P����)</td>")
			else
				Response.Write("<td><input id='Subject1' type='checkbox' name='Subject1' > ���@�ߦ۲߯Z(�P���@ �� �P����)</td>")
			end if 
		elseif gradeyear=2 then
			if identity="pro" then
				Response.Write("<td><input id='Subject1' type='checkbox' name='Subject1' > ��O�ߦ۲߯Z(�P���@ �� �P����)</td>")
			else
				Response.Write("<td><input id='Subject1' type='checkbox' name='Subject1' > ���G�ߦ۲߯Z(�P���@ �� �P����)</td>")
			end if 
		elseif gradeyear=3 and left(classname,1) <> "��" then 
				Response.Write("<td><input id='Subject1' type='checkbox' name='Subject1' > ���T�ߦ۲߯Z(�P���@ �� �P����)</td>")			
		elseif gradeyear=3 and left(classname,1) = "��"  then
				Response.Write("<td><input id='Subject1' type='checkbox' name='Subject1' disabled > �z���ݶ�g </td>")
		end if 
		
		' Table2 ��� : �A�X�ѥ[��H 
		if identity="pro" then
			Response.Write("<td><p align='left'><font color='#000000'><span style='font-size: 12pt'> �@ </span></font></td>")
		else
			Response.Write("<td><p align='left'><font color='#000000'><span style='font-size: 12pt'> ����ǥ� </span></font></td>")
		end if 
		Response.Write("</tr>")
	end if 'course1
	' --------------------------- course2	�ƾ� ���@
	if course2=true or course2=false then
		if gradeyear=1 then
			Response.Write("<tr>")
			if identity="pro" then
				Response.Write("<td><input id='Subject2' type='checkbox' name='Subject2'> ��O�ƾǩ]���Z</td>")
			elseif left(classname,1)="��" then 
				if right(classname,1)="��" then 			
					Response.Write("<td><input id='Subject2' type='checkbox' name='Subject2'> ���@�Ҽƾǥ[�j�Z(�P���G)</td>")					
				elseif right(classname,1)="�A" then 
					Response.Write("<td><input id='Subject2' type='checkbox' name='Subject2'> ���@�A�ƾǥ[�j�Z(�P���G)</td>")
				elseif right(classname,1)="��" then 
					Response.Write("<td><input id='Subject2' type='checkbox' name='Subject2'> ���@���ƾǥ[�j�Z(�P���|)</td>")
				end if 
			elseif left(classname,1)="��" or left(classname,1)="��" then
				if  flag2="�B��" and course2=false then
					Response.Write("<td><input id='Subject2' type='checkbox' name='Subject2' disabled> ���@��¾�a�����ƾ� A�Z(�P���@)</td>")
				else
					Response.Write("<td><input id='Subject2' type='checkbox' name='Subject2'> ���@��¾�a�����ƾ� A�Z(�P���@)</td>")
				end if
			else
				if  flag1="�B��" and course2=false then
					Response.Write("<td><input id='Subject2' type='checkbox' name='Subject2' disabled> ���@��¾�ӷ~���ƾ� B�Z(�P���@)</td>")
				else
					Response.Write("<td><input id='Subject2' type='checkbox' name='Subject2'> ���@��¾�ӷ~���ƾ� B�Z(�P���@)</td>")
				end if
			end if 
			
			Response.Write("<td><p align='left'><font color='#000000'><span style='font-size: 12pt'> �@ </span></font></td>")
			Response.Write("</tr>")
		' ------------- course2	�ƾ� ���G
		elseif gradeyear=2 then
			'Response.Write("<tr>")
			'Response.Write("<td>")
			if identity="pro" then
				Response.Write("<tr>")
				Response.Write("<td><input id='Subject2' type='checkbox' name='Subject2' > ��O�ƾǩ]���Z</td>")
				Response.Write("<td><p align='left'><font color='#000000'><span style='font-size: 12pt'> �@ </span></font></td>")
				Response.Write("</tr>")
			elseif left(classname,1)="��" then
				Response.Write("<tr>")
				if right(classname,1)="��" then 				
					Response.Write("<td><input id='Subject2' type='checkbox' name='Subject2' > ���G�Ҽƾǥ[�j�Z(�P���G)</td>")
				elseif right(classname,1)="�A" then 
					Response.Write("<td><input id='Subject2' type='checkbox' name='Subject2' > ���G�A�ƾǥ[�j�Z(�P���G)</td>")
				elseif right(classname,1)="��" then 
					Response.Write("<td><input id='Subject2' type='checkbox' name='Subject2' > ���G���ƾǥ[�j�Z(�P����)</td>")			
				end if
				Response.Write("<td><p align='left'><font color='#000000'><span style='font-size: 12pt'> �@ </span></font></td>")
				Response.Write("</tr>")	
				 
			elseif left(classname,1)="��" or left(classname,1)="��" then 
				'Response.Write("<tr>")
				'Response.Write("<td><input id='Subject2' type='checkbox' name='Subject2' > ���G��¾�a�����ƾ� A�Z(�P���T)</td>")
				'Response.Write("<td><p align='left'><font color='#000000'><span style='font-size: 12pt'> �@ </span></font></td>")
				'Response.Write("</tr>")
			else
				'Response.Write("<tr>")
				'Response.Write("<td><input id='Subject2' type='checkbox' name='Subject2' > ���G��¾�ӷ~���ƾ� B�Z(�P���T)</td>")
				'Response.Write("<td><p align='left'><font color='#000000'><span style='font-size: 12pt'> �@ </span></font></td>")
				'Response.Write("</tr>")
			end if 
			
		end if 
	end if  'course2
	' --------------------------- course3	�^�� ���@
	if course3=true or course3=false then
		if gradeyear=1 then
			Response.Write("<tr>")
			if identity="pro" then
				Response.Write("<td><input id='Subject3' type='checkbox' name='Subject3' > ��O�^��]���Z</td>")
			elseif left(classname,1)="��" then 
				if right(classname,1)="��" then 
					Response.Write("<td><input id='Subject3' type='checkbox' name='Subject3' > ���@�ҭ^��]���Z(�P����)</td>")
				elseif right(classname,1)="�A" then 
					Response.Write("<td><input id='Subject3' type='checkbox' name='Subject3' > ���@�A�^��]���Z(�P����)</td>")
				elseif right(classname,1)="��" then 
					Response.Write("<td><input id='Subject3' type='checkbox' name='Subject3' > ���@���^��]���Z(�P����)</td>")
				end if 			
			else
				if flag3="�B��" and course3 = false then
					Response.Write("<td><input id='Subject3' type='checkbox' name='Subject3' disabled> ���@��¾�^��[�j�Z(�P���G)�@�@<strong>�m" & flag3 & "�n</strong></td>")
				else
					Response.Write("<td><input id='Subject3' type='checkbox' name='Subject3'> ���@��¾�^��[�j�Z(�P���G)</td>")
				end if
			end if
			Response.Write("<td><p align='left'><font color='#000000'><span style='font-size: 12pt'> �@ </span></font></td>")	
			Response.Write("</tr>")
		' --------- course3	�^�� ���G	
		elseif gradeyear=2 then
			if identity="pro" then
				Response.Write("<tr>")
				Response.Write("<td><input id='Subject3' type='checkbox' name='Subject3' > ��O�^��]���Z</td>")
				Response.Write("<td><p align='left'><font color='#000000'><span style='font-size: 12pt'> �@ </span></font></td>")
				Response.Write("</tr>")
			elseif left(classname,1)="��" then
				Response.Write("<tr>")
				if right(classname,1)="��" then 
					Response.Write("<td><input id='Subject3' type='checkbox' name='Subject3' > ���G�ҭ^��]���Z(�P���T)</td>")
				elseif right(classname,1)="�A" then 
					Response.Write("<td><input id='Subject3' type='checkbox' name='Subject3' > ���G�A�^��]���Z(�P���T)</td>")
				elseif right(classname,1)="��" then 
					Response.Write("<td><input id='Subject3' type='checkbox' name='Subject3' > ���G���^��]���Z(�P���T)</td>")
				end if 
				
				Response.Write("<td><p align='left'><font color='#000000'><span style='font-size: 12pt'> �@ </span></font></td>")
				Response.Write("</tr>")
			else
				'Response.Write("<tr>")
				'Response.Write("<td><input id='Subject3' type='checkbox' name='Subject3' > ���G��¾�^��Ʋ߯Z(�P���|)</td>")		
				'Response.Write("<td><p align='left'><font color='#000000'><span style='font-size: 12pt'> �@ </span></font></td>")
				'Response.Write("</tr>")
			end if 
		end if 
	end if  'course3
	' --------------------------- course4		
	if course4=true or course4=false then
		if gradeyear=3 and (left(classname,1)="��" or left(classname,1)="�T" or left(classname,1)="��") and identity<>"pro" then 
			
			'Response.Write("<tr>")
			
			'Response.Write("<td><input id='Subject4' type='checkbox' name='Subject4'> �|�p�`�Ʋ߯Z(�P���|)</td>")
			
			'Response.Write("<td><p align='left'><font color='#000000'><span style='font-size: 12pt'> �ӶT��� </span></font></td>")
			'Response.Write("</tr>")
		end if 
	end if
	' --------------------------- course5
	if course5=true or course5=false then
		if (gradeyear=1 or gradeyear=2 or gradeyear=3) and (left(classname,1)="�s" or left(classname,1)="�e") and identity<>"pro" then 
			
			'Response.Write("<tr>")
			
			'Response.Write("<td><input id='Subject5' type='checkbox' name='Subject5'> ��¦�ϾǯZ(�P���|)</td>")
			
			'Response.Write("<td><p align='left'><font color='#000000'><span style='font-size: 12pt'> �s�]�B�h�C���ǥ� </span></font></td>")
			'Response.Write("</tr>")
		end if 
	end if	
	' --------------------------- course6
	if course6=true or course6=false then
		if (gradeyear=1 or gradeyear=2 or gradeyear=3) and (left(classname,1)="�s" or left(classname,1)="�e") and identity<>"pro" then 
			'Response.Write("<tr>")
			'Response.Write("<td><input id='Subject6' type='checkbox' name='Subject6' > ���y�Z(�P����)</td>")
			'Response.Write("<td><p align='left'><font color='#000000'><span style='font-size: 12pt'> �s�]�B�h�C���ǥ� </span></font></td>")
			'Response.Write("</tr>")
		end if 
	end if
	' --------------------------- course7
	if course7=true or course7=false then
		
	end if
	
	Response.Write("</table>")
	
	' --------------------------- ��q , busval : BusCode ---------------------------
	Response.Write("<br><td> ��q��ܡG�]�п�ܡ^�@�@�@�@�@�@�@�@�@<a href='10102�]�����ɮը����u��.xls' target='_blank'>�U�� �]���ը����u���W�@����</a></td>")
	
	Response.Write("<table border='1' width='100%' ID='Table3'>")
	Response.Write("<tr>")
	'busval : BusCode
	'Bus_Stop1() : ����='1'
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode0' type='radio' value='0' name='Bus'  onclick='Bus_Stop0();'> �ۦ汵�e</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode1' type='radio' value='1' name='Bus'  onclick='Bus_Stop1();'> 1����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode2' type='radio' value='2' name='Bus'  onclick='Bus_Stop6();'> 6����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode3' type='radio' value='3' name='Bus'  onclick='Bus_Stop6B();'> 6B����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode4' type='radio' value='4' name='Bus'  onclick='Bus_Stop6C();'> 6C����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode5' type='radio' value='5' name='Bus'  onclick='Bus_Stop8();'> 8����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode6' type='radio' value='6' name='Bus'  onclick='Bus_Stop9();'> 9����</td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode7' type='radio' value='7' name='Bus'  onclick='Bus_Stop11();'> 11����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode8' type='radio' value='8' name='Bus'  onclick='Bus_Stop11B();'> 11B����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode9' type='radio' value='9' name='Bus'  onclick='Bus_Stop15();'> 15����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode10' type='radio' value='10' name='Bus'  onclick='Bus_Stop16();'> 16����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode11' type='radio' value='11' name='Bus'  onclick='Bus_Stop20();'> 20����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode13' type='radio' value='13' name='Bus'  onclick='Bus_Stop30();'> 30����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode14' type='radio' value='14' name='Bus'  onclick='Bus_Stop33();'> 33����</td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode15' type='radio' value='15' name='Bus'  onclick='Bus_Stop36();'> 36����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode16' type='radio' value='16' name='Bus'  onclick='Bus_Stop39();'> 39����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode17' type='radio' value='17' name='Bus'  onclick='Bus_Stop39B();'> 39B����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode18' type='radio' value='18' name='Bus'  onclick='Bus_Stop39C();'> 39C����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode19' type='radio' value='19' name='Bus'  onclick='Bus_Stop42();'> 42����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode20' type='radio' value='20' name='Bus'  onclick='Bus_Stop60();'> 60����</td>")
	Response.Write("<td bgcolor=#FBFAFA><input id='BusCode12' type='Hidden' value='12' name='Bus'  onclick='Bus_Stop20B();' disabled  > </td>")
	'Response.Write("<td bgcolor=#FBFAFA><input id='BusCode12' type='radio' value='12' name='Bus'  onclick='Bus_Stop20B();' disabled > 20B����</td>")
	Response.Write("</tr>")
	Response.Write("</table>")
	
	' --------------------------- Select ��ܨ��� ---------------------------
	Response.Write("<tr>")
	Response.Write("<br><td> ���W��ܡG�]�п�ܡ^</td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	response.write("<SELECT id='Select1' name='Select1'>")
	response.write("<option>�п�ܤW�U�����W</option>")
	if busval>0 then
		while not rs_bus.EOF
			if rs_bus("�N�X")=busid then
				flag_select="selected"
			else
				flag_select=""
			end if
			response.write("<option value=" & rs_bus("�N�X") & " " & flag_select & ">" & rs_bus("�N�X") & " " & rs_bus("���W") & "</option>")
			rs_bus.MoveNext
		wend
	end if
	response.write("</SELECT>")
	Response.Write("</tr>")
	
    Response.Write("<tr>")
    Response.Write("<br><p align='center'><td width='50%' height='60'><input id='nextsubmit'  class='button' type='submit' value='�U�@�B' name='nextsubmit' onclick='return checkvalue()' >")
    Response.Write("</td>")
    Response.Write("</tr>")
    
	Response.Write("</form>")
	Response.Write("</table>")
	Response.Write("</body>")
%> 

<script type="text/javascript">
$(document).ready(function() {
	//$("#Subject2").attr("checked", true);
	//alert($("#Subject2").attr("checked") == 'checked');
	//alert("ready");
	
	//
	// initialization
	//
	initialAddClass();
	initialCheckBox();
	
	
	//
	// event
	//
	$("#nextsubmit").click(function(event) {
	  //alert(event.pageX);
	  confirm(event);
	});
	
	$("#Subject1").click(function(event) {
	  //alert("ssssssss");
	  
	});
	$("#Subject2").click(function(event) {
	  //alert("sssssss22s");
	  
	});
	
	
	function initialAddClass()
	{
		//$(".table tr:odd").addClass("odd");
		//$("#Table2 tr:even").addClass("even");
		$('tr:even').removeClass();
		$('tr:odd').removeClass();
		//$('tr:even').addClass('even');
		//$('tr:odd').addClass('odd');
		$('#Table2 tr:even').addClass('even');
		$('#Table2 tr:odd').addClass('odd');
		//$('#Table2 tr:eq(0)').addClass('eq');
	}
	
	function initialCheckBox()
	{
		if($("#Subject1").attr("checked") == 'checked')
		{
			//$("#Subject1").attr("checked", true);
			//alert("check");
		}
		
		<%
		'sql="Select * from PermRec INNER JOIN Login on RefID=PermRecID INNER JOIN Class on PermRec.ClassID=Class.ClassID where SNum='" & Account & "' And PermRec.Active=1"
		'set rs=conn.Execute(sql)
		'gradeyear=rs("ClassYear")
		'classname=rs("ClassName")
		'course1=rs("Course1")
		'course2=rs("Course2")
		'course3=rs("Course3")
		'course4=rs("Course4")
		'course5=rs("Course5")
		'course6=rs("Course6")
		'course7=rs("Course7")
		'busval=rs("Bus")
		'busid=rs("BusStopID")
		'if isnull(rs("Password1")) then
			'identity=""
		'else
			'identity=rs("Password1")
		'end	if
		
		if course1=true then
			%>  $("#Subject1").attr("checked", true);  <%
		end if
		if course2=true then
			%>  $("#Subject2").attr("checked", true);  <%
		end if
		if course3=true then
			%>  $("#Subject3").attr("checked", true);  <%
		end if
		if course4=true then
			%>  $("#Subject4").attr("checked", true);  <%
		end if
		if course5=true then
			%>  $("#Subject5").attr("checked", true);  <%
		end if
		if course6=true then
			%>  $("#Subject6").attr("checked", true);  <%
		end if
		if course7=true then
			%>  $("#Subject7").attr("checked", true);  <%
		end if
		
		'---------- ��q ---------- 
		if busval=0 then
			%>  $("#BusCode0").attr("checked", true);  <%
		end if
		if busval=1 then
			%>  $("#BusCode1").attr("checked", true);  <%
		end if
		if busval=2 then
			%>  $("#BusCode2").attr("checked", true);  <%
		end if
		if busval=3 then
			%>  $("#BusCode3").attr("checked", true);  <%
		end if
		if busval=4 then
			%>  $("#BusCode4").attr("checked", true);  <%
		end if
		if busval=5 then
			%>  $("#BusCode5").attr("checked", true);  <%
		end if
		if busval=6 then
			%>  $("#BusCode6").attr("checked", true);  <%
		end if
		if busval=7 then
			%>  $("#BusCode7").attr("checked", true);  <%
		end if
		if busval=8 then
			%>  $("#BusCode8").attr("checked", true);  <%
		end if
		if busval=9 then
			%>  $("#BusCode9").attr("checked", true);  <%
		end if
		if busval=10 then
			%>  $("#BusCode10").attr("checked", true);  <%
		end if
		if busval=11 then
			%>  $("#BusCode11").attr("checked", true);  <%
		end if
		if busval=12 then
			%>  $("#BusCode12").attr("checked", true);  <%
		end if
		if busval=13 then
			%>  $("#BusCode13").attr("checked", true);  <%
		end if
		if busval=14 then
			%>  $("#BusCode14").attr("checked", true);  <%
		end if
		if busval=15 then
			%>  $("#BusCode15").attr("checked", true);  <%
		end if
		if busval=16 then
			%>  $("#BusCode16").attr("checked", true);  <%
		end if
		if busval=17 then
			%>  $("#BusCode17").attr("checked", true);  <%
		end if
		if busval=18 then
			%>  $("#BusCode18").attr("checked", true);  <%
		end if
		if busval=19 then
			%>  $("#BusCode19").attr("checked", true);  <%
		end if
		if busval=20 then
			%>  $("#BusCode20").attr("checked", true);  <%
		end if
		
		%>
		
	}
	
	function confirm(event)
	{
		//alert("confirm");
	}
	
});
</script>
		
</html>
