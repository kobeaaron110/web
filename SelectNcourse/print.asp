<% Session.codepage=65001%>
<% Response.CharSet="utf-8" %>
<!--#InClude File="sql_conn.asp" -->
<!-- 來源
\\203.68.204.6\e$\程式\101年\2月\1010202\SelectNcourse
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
		<!--<title>選擇夜輔暨晚自習參加科目</title> -->
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
	'------- course_cnt:修課個數, total:總價錢
	course_cnt=0
	if course1=true then
		course_cnt=course_cnt+1
	end if
	if course2=true then
		if gradeyear=1 and left(classname,1)<>"普" then
		else
			course_cnt=course_cnt+1
		end if
	end if
	if course3=true then
		if gradeyear=1 and left(classname,1)<>"普" then
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
	elseif left(classname,1)="普" and gradeyear<>3 then 
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
    if Request.Form("printsubmit")="列印" then
		'Response.Write("列印="&Request.Form("printsubmit"))
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
    '1.Noprint 區間的資料不會被列印<br>    
    '2.網頁顯示時並不會分頁，只有在列印時分頁<br>    
    '------------------------------   
    if (course1=true or course2=true or course3=true or course4=true or course5=true or course6=true or course7=true) and (bus>0 and bus<=20) and busstopid<>"" then  
		Response.Write("<tr>")
		Response.Write("<td><p align='center'><font color=purple size=4>請按此按鈕列印同意書--------></font>")
		Response.Write("<input type='button' style='font-size: 16 pt; color:#FF0000; font-weight:bold' onclick='Danger_type();window.print();' value='列印' name='printsubmit'>")
		Response.Write("</td>")
		Response.Write("</tr>")
		flag=1
    elseif (course1=true or course2=true or course3=true or course4=true or course5=true or course6=true or course7=true) and (bus=0) then  
		Response.Write("<tr>")
		Response.Write("<td><p align='center'><font color=purple size=4>請按此按鈕列印同意書--------></font>")
		Response.Write("<input type='button' style='font-size: 16 pt; color:#FF0000; font-weight:bold' onclick='Danger_type();window.print();' value='列印' name='printsubmit'>")
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
	'Response.Write("<td><p align='center'><span style='font-size: 18pt; font-family: 標楷體'> " & rs1("SysVarValue") & "學年度第" & rs2("SysVarValue") & "學期夜間輔導暨晚自習活動<strong>家長同意書</strong></span></td>")
	Response.Write("<td><p align='center'><span style='font-size: 18pt; font-family: 標楷體'> " & rs1("SchoolYear") & "學年度第" & rs1("Semester") & "學期夜間輔導暨晚自習活動<strong>家長同意書</strong></span></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><p align='left'><span style='font-size: 16pt; font-family: 標楷體'> 本人同意<u>　" & rs("ClassName") & "　</u>班 學生<u>　" & rs("Name") & "　</u>學號<u>　" & Account & "　</u>，參加學校辦理之夜間輔導暨晚自習活動，並遵守相關規定，特此證明。</span></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	Response.Write("<br><td><span style='font-size: 16pt; font-family: 標楷體'> 參加科目：</span></td>")
	Response.Write("<table border='1' width='100%' ID='Table2'>")
	Response.Write("<tr>")
	Response.Write("<td width='60% >")
	
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 16pt; font-family: 標楷體'>開　課　班　別</span></font>")
	Response.Write("</td>")
	
	Response.Write("<td width='40% >")
	
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 16pt; font-family: 標楷體'>適合參加對象</span></font>")
	Response.Write("</td>")
	Response.Write("</tr>")

  '------------------ course1 自習
  if course1=true then
	Response.Write("<tr>")
	Response.Write("<td>")
	
	
	if gradeyear=1 then 
		if identity="pro" then
			if course1=true then
				Response.Write("<input type='checkbox' name='Subject1' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 國保晚自習班(星期一 ∼ 星期五)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject1' disabled><span style='font-size: 14pt; font-family: 標楷體'> 國保晚自習班(星期一 ∼ 星期五)</span></td>")
			end if 
		else
			if course1=true then
				Response.Write("<input type='checkbox' name='Subject1' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高一晚自習班(星期一 ∼ 星期五)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject1' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高一晚自習班(星期一 ∼ 星期五)</span></td>")
			end if 
		end if 
	elseif gradeyear=2 then
		if course1=true then
			Response.Write("<input type='checkbox' name='Subject1' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高二晚自習班(星期一 ∼ 星期五)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject1' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高二晚自習班(星期一 ∼ 星期五)</span></td>")
		end if 
	else
		if course1=true then
			Response.Write("<input type='checkbox' name='Subject1' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高三晚自習班(星期一 ∼ 星期五)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject1' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高三晚自習班(星期一 ∼ 星期五)</span></td>")
		end if 
	end if 
	Response.Write("<td>")
	if identity="pro" then
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt'> 　 </span></font></td>")
	else
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 全體學生 </span></font></td>")
	end if 
	Response.Write("</tr>")
  end if 
  
  '------------------ course2 數學	
  if course2=true then
	Response.Write("<tr>")
	Response.Write("<td>")
	if gradeyear=1 then 
		if identity="pro" then
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 國保數學夜輔班</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 國保數學夜輔班</span></td>")
			end if 
		elseif left(classname,1)="普" then 
			if right(classname,1)="甲" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普一甲數學加強班(星期二)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普一甲數學加強班(星期一)</span></td>")
				end if 
			elseif right(classname,1)="乙" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普一乙數學加強班(星期三)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普一乙數學加強班(星期二)</span></td>")
				end if 
			elseif right(classname,1)="丙" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普一丙數學加強班(星期三)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普一乙數學加強班(星期三)</span></td>")
				end if 
			end if 
		elseif left(classname,1)="幼" or left(classname,1)="美" then 
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高一高職家事類數學 A班(星期一)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高一高職家事類數學 A班(星期三)</span></td>")
			end if 
		else
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高一高職商業類數學 B班(星期一)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高一高職商業類數學 B班(星期三)</span></td>")
			end if 
		end if 
	elseif gradeyear=2 then
		if left(classname,1)="幼" or left(classname,1)="美" then 
		else
		end if 
		if identity="pro" then
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 國保數學夜輔班(星期三)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 國保數學夜輔班</span></td>")
			end if 
		elseif left(classname,1)="普" then 
			if right(classname,1)="甲" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普二甲數學加強班(星期二)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普二甲數學加強班(星期三)</span></td>")
				end if 
			elseif right(classname,1)="乙" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普二乙數學加強班(星期五)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普二乙數學加強班(星期四)</span></td>")
				end if 
			elseif right(classname,1)="丙" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普二丙數學加強班(星期三)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普二丙數學加強班(星期四)</span></td>")
				end if 
			end if 
		elseif left(classname,1)="幼" or left(classname,1)="美" then 
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高二高職家事類數學 A班(星期三)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高二高職家事類數學 A班(星期四)</span></td>")
			end if 
		else
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高二高職商業類數學 B班(星期三)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高二高職商業類數學 B班(星期二)</span></td>")
			end if 
		end if 
	end if 
	Response.Write("<td>") 
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 　 </span></font></td>")
	Response.Write("</tr>")
  end if 
		
  if course3=true then
	Response.Write("<tr>")
	Response.Write("<td>")
	if gradeyear=1 then 
		if identity="pro" then
			if course3=true then
				Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 國保英文夜輔班</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 國保英文夜輔班</span></td>")
			end if 
			Response.Write("<td>")
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt'> 　 </span></font></td>")
			Response.Write("</tr>")
		elseif left(classname,1)="普" then 
			if right(classname,1)="甲" then 
				if course3=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普一甲英文加強班(星期四)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普一甲英文加強班(星期二)</span></td>")
				end if 
			elseif right(classname,1)="乙" then 
				if course3=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普一乙英文加強班(星期二)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普一乙英文加強班(星期五)</span></td>")
				end if 
			elseif right(classname,1)="丙" then 
				if course3=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普一丙英文加強班(星期一)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普一丙國英加強班(星期一)</span></td>")
				end if 
			end if 
			Response.Write("<td>")
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 　 </span></font></td>")
			Response.Write("</tr>")
		else
			if course3=true then
				Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高一高職英文加強班(星期二)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高一高職英文加強班(星期一)</span></td>")
			end if 
			Response.Write("<td>")
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 　 </span></font></td>")
			Response.Write("</tr>")
		end if 
	elseif gradeyear=2 then
		if identity="pro" then
			if course3=true then
				Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 國保英文夜輔班</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 國保英文夜輔班</span></td>")
			end if 
			Response.Write("<td>")
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt'> 　 </span></font></td>")
			Response.Write("</tr>")
		elseif left(classname,1)="普" then 
			if right(classname,1)="甲" then 
				if course3=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普二甲國英加強班(星期三)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普二甲英文加強班(星期二)</span></td>")
				end if 
			elseif right(classname,1)="乙" then 
				if course3=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普二乙國英加強班(星期一)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普二乙英文加強班(星期二)</span></td>")
				end if 
			elseif right(classname,1)="丙" then 
				if course3=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普二丙國英加強班(星期一、五)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普二丙英文加強班(星期二)</span></td>")
				end if 
			end if 
			Response.Write("<td>")
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 　 </span></font></td>")
			Response.Write("</tr>")
		else
			if course3=true then
				Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高二高職英文複習班(星期四)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高二高職英文複習班(星期四)</span></td>")
			end if 
			Response.Write("<td>") 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 　 </span></font></td>")
			Response.Write("</tr>")
		end if 
	end if 
  end if 
  if course4=true then
	if left(classname,1)="商" or left(classname,1)="貿" or left(classname,1)="資" then
		Response.Write("<tr>")
		Response.Write("<td>")
		if course4=true then
			Response.Write("<input type='checkbox' name='Subject7' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 會計總複習班(星期四)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject7' disabled><span style='font-size: 14pt; font-family: 標楷體'> 會計總複習班(星期四)</span></td>")
		end if 
		Response.Write("<td>") 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 商貿資科 </span></font></td>")
		Response.Write("</tr>")
	end if 
  end if 

	
  if course5=true then
	if (gradeyear=1 or gradeyear=2 or gradeyear=3) and (left(classname,1)="廣" or left(classname,1)="畫") and identity<>"pro" then 
		i=i+1
		Response.Write("<tr>")
		Response.Write("<td>")
		if course5=true then
			Response.Write("<input type='checkbox' name='Subject5' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 基礎圖學班(星期四)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject5' disabled><span style='font-size: 14pt; font-family: 標楷體'> 基礎素描班(星期四)</span></td>")
		end if 
		Response.Write("<td>")
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 廣設、媒體科學生 </span></font></td>")
		Response.Write("</tr>")
	end if 
  end if 
	
  if course6=true then
	if (gradeyear=1 or gradeyear=2 or gradeyear=3) and (left(classname,1)="廣" or left(classname,1)="畫") and identity<>"pro" then 
		i=i+1
		Response.Write("<tr>")
		Response.Write("<td>")
		if course6=true then
			Response.Write("<input type='checkbox' name='Subject6' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 素描班(星期五)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject6' disabled><span style='font-size: 14pt; font-family: 標楷體'> 進階素描班(星期四)</span></td>")
		end if 
		Response.Write("<td>")
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 廣設、媒體科學生 </span></font></td>")
		Response.Write("</tr>")
	
	end if 
  end if 
		
	Response.Write("</table>")
	
	'---------------------------- 交通選擇 -----------------------------
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<br><td><span style='font-size: 16pt; font-family: 標楷體'> 交通選擇：</span></td>")
	Response.Write("<table border='0' width='100%' ID='Table3'>")
	Response.Write("<tr>")
	if bus=0 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　自行接送　</u>。</span></td>")
		Response.Write("</tr>")
	else	
		sql_Bus="Select * from 輔導課校車 where BusCode=" & bus
		set rs_Bus=conn.Execute(sql_Bus)
	
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　" & rs_Bus("BusNum") & "號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
		Response.Write("<br><td bgcolor=#FBFAFA><span style='font-size: 16pt; font-family: 標楷體'> 搭乘校車代碼：<u>　" & busstopid & "　</u>。</span></td>")
		Response.Write("</tr>")
	end if
	Response.Write("</table>")
	
	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><p align='left'><span style='font-size: 16pt; font-family: 標楷體'> 夜間輔導暨晚自習費用：<u>　" & total & "　</u>元。</span></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")

	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><p align='right'><span style='font-size: 16pt; font-family: 標楷體'> 家長簽章：________________</span></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	if tel="" or isnull(tel) then
		Response.Write("<td><p align='right'><span style='font-size: 16pt; font-family: 標楷體'> 聯絡電話：________________</span></td>")
	else
		Response.Write("<td><p align='right'><span style='font-size: 16pt; font-family: 標楷體'> 聯絡電話：<u>　" & tel & "_____</u></span></td>")
	end if
	if mobile="" or isnull(mobile) then
		Response.Write("<td><p align='right'><span style='font-size: 16pt; font-family: 標楷體'> 行動電話：________________</span></td>")
	else
		Response.Write("<td><p align='right'><span style='font-size: 16pt; font-family: 標楷體'> 行動電話：<u>　" & mobile & "____</u></span></td>")
	end if
	Response.Write("</tr>")
	Response.Write("<tr><br><br><br><br><br><br></tr>")
	Response.Write("<tr>")
	Response.Write("<td><p align='center'><span style='font-size: 16pt; font-family: 標楷體'> 中　　華　　民　　國　102　年　1　月　　　　日</span></td>")
	Response.Write("</tr>")
    
	Response.Write("</form>")
	Response.Write("</table>")
elseif flag=0 then
	Response.Write("<td><span style='font-size: 18pt; color:#FF0000; font-family: 標楷體'><strong>未選取任何科目，不須列印！</strong></span><br><br></td>")
elseif flag=-1 then
	Response.Write("<td><span style='font-size: 18pt; color:#FF0000; font-family: 標楷體'><strong>校車上下車站名未選取，請先選擇後再列印！</strong></span><br><br></td>")
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
		
		'---------- 交通 ---------- 
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
