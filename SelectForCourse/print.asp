<% Session.codepage=65001%>
<% Response.CharSet="utf-8" %>
<!--#InClude File="sql_conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=big5" />
		<!--<title>��ܩ]���[�ߦ۲߰ѥ[���</title> -->
		<title>Select</title>
	</head>
<script language="javascript"> 
function changeRadioForeColor(){ 
        alert(Print.Subject1.style.color);
        Print.Subject1.style.color = "blue";
        alert(Print.Subject1.style.color);
} 
function Danger_type()
{
	if (document.Print.Subject1!=undefined)
		if (document.Print.Subject1.checked)
			document.Print.Subject1.disabled = false;
	if (document.Print.Subject2!=undefined)
		if (document.Print.Subject2.checked)
			document.Print.Subject2.disabled = false;
	if (document.Print.Subject3!=undefined)
		if (document.Print.Subject3.checked)
			document.Print.Subject3.disabled = false;
	if (document.Print.Subject4!=undefined)
		if (document.Print.Subject4.checked)
			document.Print.Subject4.disabled = false;
	if (document.Print.Subject5!=undefined)
		if (document.Print.Subject5.checked)
			document.Print.Subject5.disabled = false;
	if (document.Print.Subject6!=undefined)
		if (document.Print.Subject6.checked)
			document.Print.Subject6.disabled = false;
	if (document.Print.Subject7!=undefined)
		if (document.Print.Subject7.checked)
			document.Print.Subject7.disabled = false;
}
</script> 

	<body bgcolor="#C7EDCC">
<Br>
<%
    Response.Write("<style media=print>")
    Response.Write(".Noprint {display: none;}")
    Response.Write(".PageNext {page-break-after: always;}")
    Response.Write("</style>")
        
	Account=Session("LoginID")
	Password=Session("Password")
	UserLevel=Session("LogonLevel")
	sql="Select * from PermRec INNER JOIN Login on RefID=PermRecID INNER JOIN Class on PermRec.ClassID=Class.ClassID where SNum='" & Account & "'"
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
	bus=rs("Bus")
	busstopid=rs("BusStopID")
	tel=rs("Tel")
	mobile=rs("Mobile")
	if isnull(rs("Password1")) then
		identity=""
	else
		identity=rs("Password1")
	end	if
	course_cnt=0
	if course1=true then
		course_cnt=course_cnt+1
	end if
	if course2=true then
		course_cnt=course_cnt+1
	end if
	if course3=true then
		course_cnt=course_cnt+1
	end if
	if course4=true then
		course_cnt=course_cnt+1
	end if
	if course5=true then
		course_cnt=course_cnt+1
	end if
	if course6=true then
		course_cnt=course_cnt+1
	end if
	if course7=true then
		course_cnt=course_cnt+1
	end if
	' pro : �N��O ��O�Z
	' course_cnt : �ߦ۲߭׽ҭӼ�
	' total : �]�����ɺ[�ߦ۲߶O��
	if identity="pro" then
		total=course_cnt*2500
	elseif left(classname,1)="��" and gradeyear<>3 then 
		total=course_cnt*1200
	elseif gradeyear=3 then
		if course_cnt=1 then
			total=1700
		elseif course_cnt=2 then
			total=3200
		elseif course_cnt>2 then
			total=4800
		end if
	else'if gradeyear=2 or gradeyear=3 then
		if course_cnt=1 then
			total=2500
		elseif course_cnt=2 then
			total=4800
		elseif course_cnt>2 then
			total=7000
		end if
	end if
    if Request.Form("printsubmit")="�C�L" then
		'Response.Write("�C�L="&Request.Form("printsubmit"))
		Response.Redirect "print.asp"
	end if
	'sql1="Select * from SystemVariable Where SysVarName = 'SelectCourseSchoolYear'"
	'set rs1=conn.Execute(sql1)
	'sql2="Select * from SystemVariable Where SysVarName = 'SelectCourseSemester'"
	'set rs2=conn.Execute(sql2)
	sql1="Select * From Condition Where ConditionID=1"
	set rs1=conn.Execute(sql1)
	
    Response.Write("<div class='Noprint'>")
    '------------------------------<br>    
    '1.Noprint�϶�����Ƥ��|�Q�C�L<br>    
    '2.������ܮɨä��|�����A�u���b�C�L�ɤ���<br>    
    '------------------------------   
    if (course1=true or course2=true or course3=true or course4=true or course5=true or course6=true or course7=true) and (bus>0 and bus<=18) and busstopid<>"" then  
		Response.Write("<tr>")
		Response.Write("<td><p align='center'><font color=purple size=4>�Ы������s�C�L�P�N��--------></font>")
		Response.Write("<input type='button' style='font-size: 16 pt; color:#FF0000; font-weight:bold' onclick='Danger_type();window.print();' value='�C�L' name='printsubmit'>")
		Response.Write("</td>")
		Response.Write("</tr>")
		flag=1
    elseif (course1=true or course2=true or course3=true or course4=true or course5=true or course6=true or course7=true) and (bus=0) then  
		Response.Write("<tr>")
		Response.Write("<td><p align='center'><font color=purple size=4>�Ы������s�C�L�P�N��--------></font>")
		Response.Write("<input type='button' style='font-size: 16 pt; color:#FF0000; font-weight:bold' onclick='Danger_type();window.print();' value='�C�L' name='printsubmit'>")
		Response.Write("</td>")
		Response.Write("</tr>")
		flag=1
    else
		if course1=false and course2=false and course3=false and course4=false and course5=false and course6=false and course7=false then
		'if (course1=true or course2=true or course3=true or course4=true or course5=true) and bus>0 then
			flag=0
		elseif (course1=true or course2=true or course3=true or course4=true or course5=true or course6=true or course7=true) and bus>0 and (busstopid="" or isnull(busstopid)) then
			flag=-1
		end if
    end if
    Response.Write("</div>")
    
if flag=1 then
	Response.Write("<table border='0' width='100%' ID='Table1'>")
	Response.Write("<form method='POST' name='Print' action='print.asp' ID='Form1'>")
    
	Response.Write("<tr>")
	'Response.Write("<td><p align='center'><span style='font-size: 18pt; font-family: �з���'> " & rs1("SysVarValue") & "�Ǧ~�ײ�" & rs2("SysVarValue") & "�Ǵ��]�����ɺ[�ߦ۲߬���<strong>�a���P�N��</strong></span></td>")
	Response.Write("<td><p align='center'><span style='font-size: 18pt; font-family: �з���'> " & rs1("SchoolYear") & "�Ǧ~�ײ�" & rs1("Semester") & "�Ǵ��]�����ɺ[�ߦ۲߬���<strong>�a���P�N��</strong></span></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><p align='left'><span style='font-size: 16pt; font-family: �з���'> ���H�P�N<u>�@" & rs("ClassName") & "�@</u>�Z �ǥ�<u>�@" & rs("Name") & "�@</u>�Ǹ�<u>�@" & Account & "�@</u>�A�ѥ[�Ǯտ�z���]�����ɺ[�ߦ۲߬��ʡA�ÿ�u�����W�w�A�S���ҩ��C</span></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	Response.Write("<br><td><span style='font-size: 16pt; font-family: �з���'> �ѥ[��ءG</span></td>")
	Response.Write("<table border='1' width='100%' ID='Table2'>")
	Response.Write("<tr>")
	i=i+1
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB width='60%'")
	else 
		Response.Write("<td bgcolor=#FBFAFA width='60%'")
	end if
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 16pt; font-family: �з���'>�}�@�ҡ@�Z�@�O</span></font>")
	Response.Write("</td>")
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB width='40%'>")
	else 
		Response.Write("<td bgcolor=#FBFAFA width='40%'>")
	end if
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 16pt; font-family: �з���'>�A�X�ѥ[��H</span></font>")
	Response.Write("</td>")
	Response.Write("</tr>")
  ' --------------------------- course �w��ܪ��ҵ{ ---------------------------
  '-------------------- course1 �� 1~3 �ߦ۲߯Z	
  if course1=true then
	Response.Write("<tr>")
	i=i+1
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB>")
	else 
		Response.Write("<td bgcolor=#FBFAFA>")
	end if 
	if gradeyear=1 then 
		if identity="pro" then
			if course1=true then
				Response.Write("<input type='checkbox' name='Subject1' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ��O�ߦ۲߯Z(�P���@ �� �P����)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject1' disabled><span style='font-size: 14pt; font-family: �з���'> ��O�ߦ۲߯Z(�P���@ �� �P����)</span></td>")
			end if 
		else
			if course1=true then
				Response.Write("<input type='checkbox' name='Subject1' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@�ߦ۲߯Z(�P���@ �� �P����)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject1' disabled><span style='font-size: 14pt; font-family: �з���'> ���@�ߦ۲߯Z(�P���@ �� �P����)</span></td>")
			end if 
		end if 
	elseif gradeyear=2 then
		if course1=true then
			Response.Write("<input type='checkbox' name='Subject1' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���G�ߦ۲߯Z(�P���@ �� �P����)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject1' disabled><span style='font-size: 14pt; font-family: �з���'> ���G�ߦ۲߯Z(�P���@ �� �P����)</span></td>")
		end if 
	else
		'if gradeyear=3 ���T
		if course1=true then
			Response.Write("<input type='checkbox' name='Subject1' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���T�ߦ۲߯Z(�P���@ �� �P����)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject1' disabled><span style='font-size: 14pt; font-family: �з���'> ���T�ߦ۲߯Z(�P���@ �� �P����)</span></td>")
		end if 
	end if 
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB>")
	else 
		Response.Write("<td bgcolor=#FBFAFA>")
	end if 
	if identity="pro" then
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt'> �@ </span></font></td>")
	else
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> ����ǥ� </span></font></td>")
	end if 
	Response.Write("</tr>")
  end if 
  
  '--------------------	course2 �ƾ�
  if course2=true then
	Response.Write("<tr>")
	i=i+1
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB>")
	else 
		Response.Write("<td bgcolor=#FBFAFA>")
	end if 
	if gradeyear=1 then 
		if identity="pro" then
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ��O�ƾǩ]���Z</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ��O�ƾǩ]���Z</span></td>")
			end if 
		elseif left(classname,1)="��" then 
			if right(classname,1)="��" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@�Ҽƾǥ[�j�Z(�P����)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ���@�Ҽƾǥ[�j�Z(�P���@)</span></td>")
				end if 
			elseif right(classname,1)="�A" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@�A�ƾǥ[�j�Z(�P����)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ���@�A�ƾǥ[�j�Z(�P���G)</span></td>")
				end if 
			elseif right(classname,1)="��" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@���ƾǥ[�j�Z(�P���T)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ���@�A�ƾǥ[�j�Z(�P���T)</span></td>")
				end if 
			end if 
		elseif left(classname,1)="��" or left(classname,1)="��" then 
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@��¾�a�����ƾ� A�Z(�P����)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ���@��¾�a�����ƾ� A�Z(�P���T)</span></td>")
			end if 
		else
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@��¾�ӷ~���ƾ� B�Z(�P���|)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ���@��¾�ӷ~���ƾ� B�Z(�P���T)</span></td>")
			end if 
		end if 
	elseif gradeyear=2 then
		if left(classname,1)="��" or left(classname,1)="��" then 
		else
		end if 
		if identity="pro" then
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ��O�ƾǩ]���Z(�P���T)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ��O�ƾǩ]���Z</span></td>")
			end if 
		elseif left(classname,1)="��" then 
			if right(classname,1)="��" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���G�ҼƾǽƲ߯Z(�P���@)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ���G�Ҽƾǥ[�j�Z(�P���T)</span></td>")
				end if 
			elseif right(classname,1)="�A" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���G�A�ƾǽƲ߯Z(�P���G)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ���G�A�ƾǥ[�j�Z(�P���|)</span></td>")
				end if 
			end if 
		elseif left(classname,1)="��" or left(classname,1)="��" then 
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���G��¾�a�����ƾ� A�Z(�P���T)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ���G��¾�a�����ƾ� A�Z(�P���|)</span></td>")
			end if 
		else
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���G��¾�ӷ~���ƾ� B�Z(�P���T)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ���G��¾�ӷ~���ƾ� B�Z(�P���G)</span></td>")
			end if 
		end if 
	else
		'if left(classname,1)="�s" then 
		'	if course2=true then
		'		Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> �Ϥ�ժ��A�ůZ(�P���@)</span></strong></font></td>")
		'	'else
		'	'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ���T��¾�a�����ƾ� A�Z(�P���@)</span></td>")
		'	end if 
		if left(classname,1)="��" or left(classname,1)="�T" or left(classname,1)="��" then 
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> �|�p�`�Ʋ߯Z(�P���@)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ���T��¾�ӷ~���ƾ� B�Z(�P���@)</span></td>")
			end if 
		end if 
	end if 
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB>")
	else 
		Response.Write("<td bgcolor=#FBFAFA>")
	end if 
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �@ </span></font></td>")
	Response.Write("</tr>")
  end if 
  
  '--------------------	course3  
  if course3=true then
	if gradeyear=1 And left(classname,1)<>"��" then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course3=true then
		'	Response.Write("<input type='checkbox' name='Subject3' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> �q���A���˩w�Z(�P���T)</span></strong></font></td>")
		'else
			Response.Write("<input type='checkbox' name='Subject3' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@��¾�^���¦�Z(�P���G)</span></strong></font></td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'>  </span></font></td>")
		Response.Write("</tr>")
	end if 
  end if 

  '--------------------	course4  
  if course4=true then
	Response.Write("<tr>")
	i=i+1
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB>")
	else 
		Response.Write("<td bgcolor=#FBFAFA>")
	end if 
	if gradeyear=1 then 
		if identity="pro" then
			if course4=true then
				Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ��O�^��]���Z</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ��O�^��]���Z</span></td>")
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt'> �@ </span></font></td>")
			Response.Write("</tr>")
		elseif left(classname,1)="��" then 
			if right(classname,1)="��" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@�Ұ�^�[�j�Z(�P���T)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ���@�ҭ^��[�j�Z(�P���G)</span></td>")
				end if 
			elseif right(classname,1)="�A" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@�A��^�[�j�Z(�P���@)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ���@�A�^��[�j�Z(�P����)</span></td>")
				end if 
			elseif right(classname,1)="��" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@����^�[�j�Z(�P���@)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ���@����^�[�j�Z(�P���@)</span></td>")
				end if 
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �@ </span></font></td>")
			Response.Write("</tr>")
		else
			if course4=true then
				Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@��¾�^��i���Z(�P���G)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ���@��¾�^��[�j�Z(�P���@)</span></td>")
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �@ </span></font></td>")
			Response.Write("</tr>")
		end if 
	elseif gradeyear=2 then
		if identity="pro" then
			if course4=true then
				Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ��O�^��]���Z</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ��O�^��]���Z</span></td>")
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt'> �@ </span></font></td>")
			Response.Write("</tr>")
		elseif left(classname,1)="��" then 
			if right(classname,1)="��" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���G�Ұ�^�[�j�Z(�P���G)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ���G�ҭ^��[�j�Z(�P���G)</span></td>")
				end if 
			elseif right(classname,1)="�A" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���G�A��^�[�j�Z(�P����)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ���G�A�^��[�j�Z(�P���G)</span></td>")
				end if 
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �@ </span></font></td>")
			Response.Write("</tr>")
		else
			if course4=true then
				Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���G��¾�^��Ʋ߯Z(�P���|)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ���G��¾�^��Ʋ߯Z(�P���|)</span></td>")
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �@ </span></font></td>")
			Response.Write("</tr>")
		end if 
	'else
	'	if left(classname,1)="��" or left(classname,1)="�T" or left(classname,1)="��" then
	'		Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> �g���`�Ʋ߯Z(�P����)</span></strong></font></td>")
	'	'else
	'	'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ���T��¾�^��[�j�Z(�P����)</span></td>")
	'	end if 
	'	if (i mod 2)=1 then 
	'		Response.Write("<td bgcolor=#EBEBEB>")
	'	else 
	'		Response.Write("<td bgcolor=#FBFAFA>")
	'	end if 
	'	Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'>  </span></font></td>")
	'	Response.Write("</tr>")
	end if 
  end if 

  '--------------------	  
  if course5=true then
	if (gradeyear=1 or gradeyear=2 or gradeyear=3) and (left(classname,1)="�s" or left(classname,1)="�e") and identity<>"pro" then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course5=true then
			Response.Write("<input type='checkbox' name='Subject5' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ��¦���y�Z(�P���@)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject5' disabled><span style='font-size: 14pt; font-family: �з���'> ��¦���y�Z(�P���|)</span></td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �s�]�B�C���ǥ� </span></font></td>")
		Response.Write("</tr>")
	end if 
  end if 

  '--------------------	  
  if course6=true then
	if (gradeyear=1 or gradeyear=2 or gradeyear=3) and (left(classname,1)="�s" or left(classname,1)="�e") and identity<>"pro" then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course6=true then
			Response.Write("<input type='checkbox' name='Subject6' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> �i�����y�Z(�P���@)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject6' disabled><span style='font-size: 14pt; font-family: �з���'> �i�����y�Z(�P���|)</span></td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �s�]�B�C���ǥ� </span></font></td>")
		Response.Write("</tr>")
	'elseif gradeyear=2 and left(classname,1)="�s" then
	'	i=i+1
	'	Response.Write("<tr>")
	'	if (i mod 2)=1 then 
	'		Response.Write("<td bgcolor=#EBEBEB>")
	'	else 
	'		Response.Write("<td bgcolor=#FBFAFA>")
	'	end if 
	'	if course6=true then
	'		Response.Write("<input type='checkbox' name='Subject6' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> �i�����y�Z(�P���|)</span></strong></font></td>")
	'	'else
	'	'	Response.Write("<input type='checkbox' name='Subject6' disabled><span style='font-size: 14pt; font-family: �з���'> �i�����y�Z(�P���|)</span></td>")
	'	end if 
	'	if (i mod 2)=1 then 
	'		Response.Write("<td bgcolor=#EBEBEB>")
	'	else 
	'		Response.Write("<td bgcolor=#FBFAFA>")
	'	end if 
	'	Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �s�]�B�C���ǥ� </span></font></td>")
	'	Response.Write("</tr>")
	end if 
  end if 

  '--------------------	  
  if course7=true then
	'if gradeyear=1 and (left(classname,1)="��" or left(classname,1)="��" or left(classname,1)="�T") and identity<>"pro" then 
	'	i=i+1
	'	Response.Write("<tr>")
	'	if (i mod 2)=1 then 
	'		Response.Write("<td bgcolor=#EBEBEB>")
	'	else 
	'		Response.Write("<td bgcolor=#FBFAFA>")
	'	end if 
	'	if course7=true then
	'		Response.Write("<input type='checkbox' name='Subject7' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@�|�p�[�j�Z(�P����)</span></strong></font></td>")
	'	'else
	'	'	Response.Write("<input type='checkbox' name='Subject7' disabled><span style='font-size: 14pt; font-family: �з���'> ���@�|�p�[�j�Z(�P����)</span></td>")
	'	end if 
	'	if (i mod 2)=1 then 
	'		Response.Write("<td bgcolor=#EBEBEB>")
	'	else 
	'		Response.Write("<td bgcolor=#FBFAFA>")
	'	end if 
	'	Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �Ӹ�T��A���[�j�|�p�� </span></font></td>")
	'	Response.Write("</tr>")
	if (gradeyear=2 or gradeyear=3) and (left(classname,1)="�s" or left(classname,1)="�e") then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course7=true then
			Response.Write("<input type='checkbox' name='Subject7' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���J����¦�Z(�P���G)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject7' disabled><span style='font-size: 14pt; font-family: �з���'> ���J���Z(�P���T)</span></td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �s�]�B�C���ǥ� </span></font></td>")
		Response.Write("</tr>")
	end if 
  end if 
		
	Response.Write("</table>")
	
	'if bus=0 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: �з���'>�ۦ汵�e</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>1����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>6����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>9����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>15����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>16����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>20����</span></td>")
	'elseif bus=1 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>�ۦ汵�e</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: �з���'>1����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>6����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>9����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>15����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>16����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>20����</span></td>")
	'elseif bus=2 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>�ۦ汵�e</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>1����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: �з���'>6����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>9����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>15����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>16����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>20����</span></td>")
	'elseif bus=3 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>�ۦ汵�e</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>1����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>6����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: �з���'>9����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>15����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>16����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>20����</span></td>")
	'elseif bus=4 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>�ۦ汵�e</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>1����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>6����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>9����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: �з���'>15����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>16����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>20����</span></td>")
	'elseif bus=5 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>�ۦ汵�e</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>1����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>6����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>9����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>15����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: �з���'>16����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>20����</span></td>")
	'elseif bus=6 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>�ۦ汵�e</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>1����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>6����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>9����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>15����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>16����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: �з���'>20����</span></td>")
	'else
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>�ۦ汵�e</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>1����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>6����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>9����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>15����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>16����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>20����</span></td>")
	'end if
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<br><td><span style='font-size: 16pt; font-family: �з���'> ��q��ܡG</span></td>")
	Response.Write("<table border='0' width='100%' ID='Table3'>")
	Response.Write("<tr>")
	if bus=0 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@�ۦ汵�e�@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=1 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@1�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=2 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@6�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=3 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@8�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=4 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@9�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=5 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@11�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=6 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@15�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=7 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@16�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=8 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@20�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=9 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@26�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=10 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@29�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=11 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@30�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=12 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@33�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=13 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@36�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=14 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@39�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=15 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@42�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=16 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@60�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=17 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@9B�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=18 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@9C�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	end if
	Response.Write("<br><td bgcolor=#FBFAFA><span style='font-size: 16pt; font-family: �з���'> �f���ը��N�X�G<u>�@" & busstopid & "�@</u>�C</span></td>")
	'if bus=7 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: �з���'>�s���ꤤ(8�H�y)</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>26����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>29����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>30����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>33����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>36����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>39����</span></td>")
	'elseif bus=8 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>�s���ꤤ(8�H�y)</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: �з���'>26����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>29����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>30����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>33����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>36����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>39����</span></td>")
	'elseif bus=9 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>�s���ꤤ(8�H�y)</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>26����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: �з���'>29����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>30����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>33����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>36����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>39����</span></td>")
	'elseif bus=10 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>�s���ꤤ(8�H�y)</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>26����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>29����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: �з���'>30����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>33����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>36����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>39����</span></td>")
	'elseif bus=11 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>�s���ꤤ(8�H�y)</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>26����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>29����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>30����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: �з���'>33����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>36����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>39����</span></td>")
	'elseif bus=12 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>�s���ꤤ(8�H�y)</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>26����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>29����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>30����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>33����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: �з���'>36����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>39����</span></td>")
	'elseif bus=13 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>�s���ꤤ(8�H�y)</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>26����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>29����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>30����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>33����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>36����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: �з���'>39����</span></td>")
	'else
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>�s���ꤤ(8�H�y)</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>26����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>29����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>30����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>33����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>36����</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>39����</span></td>")
	'end if
	Response.Write("</tr>")
	Response.Write("<tr>")
	'if bus=14 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: �з���'>39����B�u</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>42����</span></td>")
	'elseif bus=15 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>39����B�u</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: �з���'>42����</span></td>")
	'else
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>39����B�u</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: �з���'>42����</span></td>")
	'end if
	Response.Write("</tr>")
	Response.Write("</table>")
	
	' --------------------------- �۲߶O�� total ---------------------------
	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><p align='left'><span style='font-size: 16pt; font-family: �з���'> �]�����ɺ[�ߦ۲߶O�ΡG<u>�@" & total & "�@</u>���C</span></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")

	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><p align='right'><span style='font-size: 16pt; font-family: �з���'> �a��ñ���G________________</span></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	if tel="" or isnull(tel) then
		Response.Write("<td><p align='right'><span style='font-size: 16pt; font-family: �з���'> �p���q�ܡG________________</span></td>")
	else
		Response.Write("<td><p align='right'><span style='font-size: 16pt; font-family: �з���'> �p���q�ܡG<u>�@" & tel & "_____</u></span></td>")
	end if
	if mobile="" or isnull(mobile) then
		Response.Write("<td><p align='right'><span style='font-size: 16pt; font-family: �з���'> ��ʹq�ܡG________________</span></td>")
	else
		Response.Write("<td><p align='right'><span style='font-size: 16pt; font-family: �з���'> ��ʹq�ܡG<u>�@" & mobile & "____</u></span></td>")
	end if
	Response.Write("</tr>")
	Response.Write("<tr><br><br><br><br><br><br></tr>")
	Response.Write("<tr>")
	Response.Write("<td><p align='center'><span style='font-size: 16pt; font-family: �з���'> ���@�@�ء@�@���@�@��@�E�@�Q�@�E�@�~�@�K�@��@�@�@�@��</span></td>")
	Response.Write("</tr>")
    
	Response.Write("</form>")
	Response.Write("</table>")
elseif flag=0 then
	Response.Write("<td><span style='font-size: 18pt; color:#FF0000; font-family: �з���'><strong>����������ءA�����C�L�I</strong></span><br><br></td>")
elseif flag=-1 then
	Response.Write("<td><span style='font-size: 18pt; color:#FF0000; font-family: �з���'><strong>�ը��W�U�����W������A�Х���ܫ�A�C�L�I</strong></span><br><br></td>")
end if

	Response.Write("</body>")
%> 
</html>
