<% Session.codepage=65001%>
<% Response.CharSet="utf-8" %>
<!--#InClude File="sql_conn.asp" -->
<!-- �ӷ�
\\203.68.204.6\e$\�{��\101�~\2��\1010202\SelectNcourse
-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
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
	total=0
	if isnull(rs("Password1")) then
		identity=""
	else
		identity=rs("Password1")
	end	if
	'------- course_cnt:�׽ҭӼ�, total:�`����
	course_cnt=0
	if course1=true then
		course_cnt=course_cnt+1
	end if
	if course2=true then
		if gradeyear=1 and left(classname,1)<>"��" then
		else
			course_cnt=course_cnt+1
		end if
	end if
	if course3=true then
		if gradeyear=1 and left(classname,1)<>"��" then
		else
			course_cnt=course_cnt+1
		end if
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
	elseif gradeyear=2 or gradeyear=1 then
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
    '1.Noprint �϶�����Ƥ��|�Q�C�L<br>    
    '2.������ܮɨä��|�����A�u���b�C�L�ɤ���<br>    
    '------------------------------   
    if (course1=true or course2=true or course3=true or course4=true or course5=true or course6=true or course7=true) and (bus>0 and bus<=20) and busstopid<>"" then  
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
	Response.Write("<td width='60% >")
	
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 16pt; font-family: �з���'>�}�@�ҡ@�Z�@�O</span></font>")
	Response.Write("</td>")
	
	Response.Write("<td width='40% >")
	
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 16pt; font-family: �з���'>�A�X�ѥ[��H</span></font>")
	Response.Write("</td>")
	Response.Write("</tr>")

  '------------------ course1 �۲�
  if course1=true then
	Response.Write("<tr>")
	Response.Write("<td>")
	
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
		if course1=true then
			Response.Write("<input type='checkbox' name='Subject1' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���T�ߦ۲߯Z(�P���@ �� �P����)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject1' disabled><span style='font-size: 14pt; font-family: �з���'> ���T�ߦ۲߯Z(�P���@ �� �P����)</span></td>")
		end if 
	end if 
	Response.Write("<td>")
	if identity="pro" then
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt'> �@ </span></font></td>")
	else
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> ����ǥ� </span></font></td>")
	end if 
	Response.Write("</tr>")
  end if 
  
  '------------------ course2 �ƾ�	
  if course2=true then
	Response.Write("<tr>")
	Response.Write("<td>")
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
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@�Ҽƾǥ[�j�Z(�P���G)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ���@�Ҽƾǥ[�j�Z(�P���@)</span></td>")
				end if 
			elseif right(classname,1)="�A" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@�A�ƾǥ[�j�Z(�P���T)</span></strong></font></td>")
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
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@��¾�a�����ƾ� A�Z(�P���@)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ���@��¾�a�����ƾ� A�Z(�P���T)</span></td>")
			end if 
		else
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@��¾�ӷ~���ƾ� B�Z(�P���@)</span></strong></font></td>")
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
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���G�Ҽƾǥ[�j�Z(�P���G)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ���G�Ҽƾǥ[�j�Z(�P���T)</span></td>")
				end if 
			elseif right(classname,1)="�A" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���G�A�ƾǥ[�j�Z(�P����)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ���G�A�ƾǥ[�j�Z(�P���|)</span></td>")
				end if 
			elseif right(classname,1)="��" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���G���ƾǥ[�j�Z(�P���T)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: �з���'> ���G���ƾǥ[�j�Z(�P���|)</span></td>")
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
	end if 
	Response.Write("<td>") 
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �@ </span></font></td>")
	Response.Write("</tr>")
  end if 
		
  if course3=true then
	Response.Write("<tr>")
	Response.Write("<td>")
	if gradeyear=1 then 
		if identity="pro" then
			if course3=true then
				Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ��O�^��]���Z</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ��O�^��]���Z</span></td>")
			end if 
			Response.Write("<td>")
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt'> �@ </span></font></td>")
			Response.Write("</tr>")
		elseif left(classname,1)="��" then 
			if right(classname,1)="��" then 
				if course3=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@�ҭ^��[�j�Z(�P���|)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ���@�ҭ^��[�j�Z(�P���G)</span></td>")
				end if 
			elseif right(classname,1)="�A" then 
				if course3=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@�A�^��[�j�Z(�P���G)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ���@�A�^��[�j�Z(�P����)</span></td>")
				end if 
			elseif right(classname,1)="��" then 
				if course3=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@���^��[�j�Z(�P���@)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ���@����^�[�j�Z(�P���@)</span></td>")
				end if 
			end if 
			Response.Write("<td>")
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �@ </span></font></td>")
			Response.Write("</tr>")
		else
			if course3=true then
				Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���@��¾�^��[�j�Z(�P���G)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ���@��¾�^��[�j�Z(�P���@)</span></td>")
			end if 
			Response.Write("<td>")
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �@ </span></font></td>")
			Response.Write("</tr>")
		end if 
	elseif gradeyear=2 then
		if identity="pro" then
			if course3=true then
				Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ��O�^��]���Z</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ��O�^��]���Z</span></td>")
			end if 
			Response.Write("<td>")
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt'> �@ </span></font></td>")
			Response.Write("</tr>")
		elseif left(classname,1)="��" then 
			if right(classname,1)="��" then 
				if course3=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���G�Ұ�^�[�j�Z(�P���T)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ���G�ҭ^��[�j�Z(�P���G)</span></td>")
				end if 
			elseif right(classname,1)="�A" then 
				if course3=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���G�A��^�[�j�Z(�P���@)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ���G�A�^��[�j�Z(�P���G)</span></td>")
				end if 
			elseif right(classname,1)="��" then 
				if course3=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���G����^�[�j�Z(�P���@�B��)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ���G���^��[�j�Z(�P���G)</span></td>")
				end if 
			end if 
			Response.Write("<td>")
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �@ </span></font></td>")
			Response.Write("</tr>")
		else
			if course3=true then
				Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���G��¾�^��Ʋ߯Z(�P���|)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: �з���'> ���G��¾�^��Ʋ߯Z(�P���|)</span></td>")
			end if 
			Response.Write("<td>") 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �@ </span></font></td>")
			Response.Write("</tr>")
		end if 
	end if 
  end if 
  if course4=true then
	if left(classname,1)="��" or left(classname,1)="�T" or left(classname,1)="��" then
		Response.Write("<tr>")
		Response.Write("<td>")
		if course4=true then
			Response.Write("<input type='checkbox' name='Subject7' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> �|�p�`�Ʋ߯Z(�P���|)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject7' disabled><span style='font-size: 14pt; font-family: �з���'> �|�p�`�Ʋ߯Z(�P���|)</span></td>")
		end if 
		Response.Write("<td>") 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �ӶT��� </span></font></td>")
		Response.Write("</tr>")
	end if 
  end if 

	
  if course5=true then
	if (gradeyear=1 or gradeyear=2 or gradeyear=3) and (left(classname,1)="�s" or left(classname,1)="�e") and identity<>"pro" then 
		i=i+1
		Response.Write("<tr>")
		Response.Write("<td>")
		if course5=true then
			Response.Write("<input type='checkbox' name='Subject5' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ��¦�ϾǯZ(�P���|)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject5' disabled><span style='font-size: 14pt; font-family: �з���'> ��¦���y�Z(�P���|)</span></td>")
		end if 
		Response.Write("<td>")
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �s�]�B�C���ǥ� </span></font></td>")
		Response.Write("</tr>")
	end if 
  end if 
	
  if course6=true then
	if (gradeyear=1 or gradeyear=2 or gradeyear=3) and (left(classname,1)="�s" or left(classname,1)="�e") and identity<>"pro" then 
		i=i+1
		Response.Write("<tr>")
		Response.Write("<td>")
		if course6=true then
			Response.Write("<input type='checkbox' name='Subject6' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: �з���'> ���y�Z(�P����)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject6' disabled><span style='font-size: 14pt; font-family: �з���'> �i�����y�Z(�P���|)</span></td>")
		end if 
		Response.Write("<td>")
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: �з���'> �s�]�B�C���ǥ� </span></font></td>")
		Response.Write("</tr>")
	
	end if 
  end if 
		
	Response.Write("</table>")
	
	'---------------------------- ��q��� -----------------------------
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<br><td><span style='font-size: 16pt; font-family: �з���'> ��q��ܡG</span></td>")
	Response.Write("<table border='0' width='100%' ID='Table3'>")
	Response.Write("<tr>")
	if bus=0 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@�ۦ汵�e�@</u>�C</span></td>")
		Response.Write("</tr>")
	else	
		sql_Bus="Select * from ���ɽҮը� where BusCode=" & bus
		set rs_Bus=conn.Execute(sql_Bus)
	
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: �з���'>�z��ܪ��O<u>�@" & rs_Bus("BusNum") & "�����@</u>�C</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
		Response.Write("<br><td bgcolor=#FBFAFA><span style='font-size: 16pt; font-family: �з���'> �f���ը��N�X�G<u>�@" & busstopid & "�@</u>�C</span></td>")
		Response.Write("</tr>")
	end if
	Response.Write("</table>")
	
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
	Response.Write("<td><p align='center'><span style='font-size: 16pt; font-family: �з���'> ���@�@�ء@�@���@�@��@102�@�~�@1�@��@�@�@�@��</span></td>")
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

<script type="text/javascript">
$(document).ready(function() {
	
	alert("print ready");
	
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
	  //confirm(event);
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
	
});
</script>
</html>
