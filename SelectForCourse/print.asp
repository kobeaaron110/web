<% Session.codepage=65001%>
<% Response.CharSet="utf-8" %>
<!--#InClude File="sql_conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=big5" />
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
	' pro : 代表是 國保班
	' course_cnt : 晚自習修課個數
	' total : 夜間輔導暨晚自習費用
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
	else'if gradeyear=2 or gradeyear=3 then
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
    '1.Noprint區間的資料不會被列印<br>    
    '2.網頁顯示時並不會分頁，只有在列印時分頁<br>    
    '------------------------------   
    if (course1=true or course2=true or course3=true or course4=true or course5=true or course6=true or course7=true) and (bus>0 and bus<=18) and busstopid<>"" then  
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
	i=i+1
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB width='60%'")
	else 
		Response.Write("<td bgcolor=#FBFAFA width='60%'")
	end if
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 16pt; font-family: 標楷體'>開　課　班　別</span></font>")
	Response.Write("</td>")
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB width='40%'>")
	else 
		Response.Write("<td bgcolor=#FBFAFA width='40%'>")
	end if
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 16pt; font-family: 標楷體'>適合參加對象</span></font>")
	Response.Write("</td>")
	Response.Write("</tr>")
  ' --------------------------- course 已選擇的課程 ---------------------------
  '-------------------- course1 高 1~3 晚自習班	
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
				Response.Write("<input type='checkbox' name='Subject1' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 國保晚自習班(星期一 ～ 星期五)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject1' disabled><span style='font-size: 14pt; font-family: 標楷體'> 國保晚自習班(星期一 ～ 星期五)</span></td>")
			end if 
		else
			if course1=true then
				Response.Write("<input type='checkbox' name='Subject1' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高一晚自習班(星期一 ～ 星期五)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject1' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高一晚自習班(星期一 ～ 星期五)</span></td>")
			end if 
		end if 
	elseif gradeyear=2 then
		if course1=true then
			Response.Write("<input type='checkbox' name='Subject1' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高二晚自習班(星期一 ～ 星期五)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject1' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高二晚自習班(星期一 ～ 星期五)</span></td>")
		end if 
	else
		'if gradeyear=3 高三
		if course1=true then
			Response.Write("<input type='checkbox' name='Subject1' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高三晚自習班(星期一 ～ 星期五)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject1' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高三晚自習班(星期一 ～ 星期五)</span></td>")
		end if 
	end if 
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB>")
	else 
		Response.Write("<td bgcolor=#FBFAFA>")
	end if 
	if identity="pro" then
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt'> 　 </span></font></td>")
	else
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 全體學生 </span></font></td>")
	end if 
	Response.Write("</tr>")
  end if 
  
  '--------------------	course2 數學
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
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 國保數學夜輔班</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 國保數學夜輔班</span></td>")
			end if 
		elseif left(classname,1)="普" then 
			if right(classname,1)="甲" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普一甲數學加強班(星期五)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普一甲數學加強班(星期一)</span></td>")
				end if 
			elseif right(classname,1)="乙" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普一乙數學加強班(星期五)</span></strong></font></td>")
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
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高一高職家事類數學 A班(星期五)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高一高職家事類數學 A班(星期三)</span></td>")
			end if 
		else
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高一高職商業類數學 B班(星期四)</span></strong></font></td>")
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
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普二甲數學複習班(星期一)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普二甲數學加強班(星期三)</span></td>")
				end if 
			elseif right(classname,1)="乙" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普二乙數學複習班(星期二)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普二乙數學加強班(星期四)</span></td>")
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
	else
		'if left(classname,1)="廣" then 
		'	if course2=true then
		'		Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 圖文組版乙級班(星期一)</span></strong></font></td>")
		'	'else
		'	'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高三高職家事類數學 A班(星期一)</span></td>")
		'	end if 
		if left(classname,1)="商" or left(classname,1)="貿" or left(classname,1)="資" then 
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 會計總複習班(星期一)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject2' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高三高職商業類數學 B班(星期一)</span></td>")
			end if 
		end if 
	end if 
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB>")
	else 
		Response.Write("<td bgcolor=#FBFAFA>")
	end if 
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 　 </span></font></td>")
	Response.Write("</tr>")
  end if 
  
  '--------------------	course3  
  if course3=true then
	if gradeyear=1 And left(classname,1)<>"普" then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course3=true then
		'	Response.Write("<input type='checkbox' name='Subject3' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 電腦乙級檢定班(星期三)</span></strong></font></td>")
		'else
			Response.Write("<input type='checkbox' name='Subject3' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高一高職英文基礎班(星期二)</span></strong></font></td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'>  </span></font></td>")
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
				Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 國保英文夜輔班</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 國保英文夜輔班</span></td>")
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt'> 　 </span></font></td>")
			Response.Write("</tr>")
		elseif left(classname,1)="普" then 
			if right(classname,1)="甲" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普一甲國英加強班(星期三)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普一甲英文加強班(星期二)</span></td>")
				end if 
			elseif right(classname,1)="乙" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普一乙國英加強班(星期一)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普一乙英文加強班(星期五)</span></td>")
				end if 
			elseif right(classname,1)="丙" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普一丙國英加強班(星期一)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普一丙國英加強班(星期一)</span></td>")
				end if 
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 　 </span></font></td>")
			Response.Write("</tr>")
		else
			if course4=true then
				Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高一高職英文進階班(星期二)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高一高職英文加強班(星期一)</span></td>")
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 　 </span></font></td>")
			Response.Write("</tr>")
		end if 
	elseif gradeyear=2 then
		if identity="pro" then
			if course4=true then
				Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 國保英文夜輔班</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 國保英文夜輔班</span></td>")
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt'> 　 </span></font></td>")
			Response.Write("</tr>")
		elseif left(classname,1)="普" then 
			if right(classname,1)="甲" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普二甲國英加強班(星期二)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普二甲英文加強班(星期二)</span></td>")
				end if 
			elseif right(classname,1)="乙" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 普二乙國英加強班(星期五)</span></strong></font></td>")
				'else
				'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 普二乙英文加強班(星期二)</span></td>")
				end if 
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 　 </span></font></td>")
			Response.Write("</tr>")
		else
			if course4=true then
				Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高二高職英文複習班(星期四)</span></strong></font></td>")
			'else
			'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高二高職英文複習班(星期四)</span></td>")
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 　 </span></font></td>")
			Response.Write("</tr>")
		end if 
	'else
	'	if left(classname,1)="商" or left(classname,1)="貿" or left(classname,1)="資" then
	'		Response.Write("<input type='checkbox' name='Subject4' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 經濟總複習班(星期五)</span></strong></font></td>")
	'	'else
	'	'	Response.Write("<input type='checkbox' name='Subject4' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高三高職英文加強班(星期五)</span></td>")
	'	end if 
	'	if (i mod 2)=1 then 
	'		Response.Write("<td bgcolor=#EBEBEB>")
	'	else 
	'		Response.Write("<td bgcolor=#FBFAFA>")
	'	end if 
	'	Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'>  </span></font></td>")
	'	Response.Write("</tr>")
	end if 
  end if 

  '--------------------	  
  if course5=true then
	if (gradeyear=1 or gradeyear=2 or gradeyear=3) and (left(classname,1)="廣" or left(classname,1)="畫") and identity<>"pro" then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course5=true then
			Response.Write("<input type='checkbox' name='Subject5' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 基礎素描班(星期一)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject5' disabled><span style='font-size: 14pt; font-family: 標楷體'> 基礎素描班(星期四)</span></td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 廣設、媒體科學生 </span></font></td>")
		Response.Write("</tr>")
	end if 
  end if 

  '--------------------	  
  if course6=true then
	if (gradeyear=1 or gradeyear=2 or gradeyear=3) and (left(classname,1)="廣" or left(classname,1)="畫") and identity<>"pro" then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course6=true then
			Response.Write("<input type='checkbox' name='Subject6' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 進階素描班(星期一)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject6' disabled><span style='font-size: 14pt; font-family: 標楷體'> 進階素描班(星期四)</span></td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 廣設、媒體科學生 </span></font></td>")
		Response.Write("</tr>")
	'elseif gradeyear=2 and left(classname,1)="廣" then
	'	i=i+1
	'	Response.Write("<tr>")
	'	if (i mod 2)=1 then 
	'		Response.Write("<td bgcolor=#EBEBEB>")
	'	else 
	'		Response.Write("<td bgcolor=#FBFAFA>")
	'	end if 
	'	if course6=true then
	'		Response.Write("<input type='checkbox' name='Subject6' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 進階素描班(星期四)</span></strong></font></td>")
	'	'else
	'	'	Response.Write("<input type='checkbox' name='Subject6' disabled><span style='font-size: 14pt; font-family: 標楷體'> 進階素描班(星期四)</span></td>")
	'	end if 
	'	if (i mod 2)=1 then 
	'		Response.Write("<td bgcolor=#EBEBEB>")
	'	else 
	'		Response.Write("<td bgcolor=#FBFAFA>")
	'	end if 
	'	Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 廣設、媒體科學生 </span></font></td>")
	'	Response.Write("</tr>")
	end if 
  end if 

  '--------------------	  
  if course7=true then
	'if gradeyear=1 and (left(classname,1)="商" or left(classname,1)="資" or left(classname,1)="貿") and identity<>"pro" then 
	'	i=i+1
	'	Response.Write("<tr>")
	'	if (i mod 2)=1 then 
	'		Response.Write("<td bgcolor=#EBEBEB>")
	'	else 
	'		Response.Write("<td bgcolor=#FBFAFA>")
	'	end if 
	'	if course7=true then
	'		Response.Write("<input type='checkbox' name='Subject7' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 高一會計加強班(星期五)</span></strong></font></td>")
	'	'else
	'	'	Response.Write("<input type='checkbox' name='Subject7' disabled><span style='font-size: 14pt; font-family: 標楷體'> 高一會計加強班(星期五)</span></td>")
	'	end if 
	'	if (i mod 2)=1 then 
	'		Response.Write("<td bgcolor=#EBEBEB>")
	'	else 
	'		Response.Write("<td bgcolor=#FBFAFA>")
	'	end if 
	'	Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 商資貿科，欲加強會計者 </span></font></td>")
	'	Response.Write("</tr>")
	if (gradeyear=2 or gradeyear=3) and (left(classname,1)="廣" or left(classname,1)="畫") then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course7=true then
			Response.Write("<input type='checkbox' name='Subject7' checked disabled><font color='blue'><strong><span style='font-size: 14pt; font-family: 標楷體'> 麥克筆基礎班(星期二)</span></strong></font></td>")
		'else
		'	Response.Write("<input type='checkbox' name='Subject7' disabled><span style='font-size: 14pt; font-family: 標楷體'> 麥克筆班(星期三)</span></td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 14pt; font-family: 標楷體'> 廣設、媒體科學生 </span></font></td>")
		Response.Write("</tr>")
	end if 
  end if 
		
	Response.Write("</table>")
	
	'if bus=0 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: 標楷體'>自行接送</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>1號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>6號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>9號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>15號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>16號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>20號車</span></td>")
	'elseif bus=1 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>自行接送</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: 標楷體'>1號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>6號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>9號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>15號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>16號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>20號車</span></td>")
	'elseif bus=2 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>自行接送</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>1號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: 標楷體'>6號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>9號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>15號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>16號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>20號車</span></td>")
	'elseif bus=3 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>自行接送</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>1號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>6號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: 標楷體'>9號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>15號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>16號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>20號車</span></td>")
	'elseif bus=4 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>自行接送</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>1號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>6號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>9號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: 標楷體'>15號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>16號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>20號車</span></td>")
	'elseif bus=5 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>自行接送</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>1號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>6號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>9號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>15號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: 標楷體'>16號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>20號車</span></td>")
	'elseif bus=6 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>自行接送</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>1號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>6號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>9號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>15號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>16號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: 標楷體'>20號車</span></td>")
	'else
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>自行接送</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>1號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>6號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>9號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>15號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>16號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>20號車</span></td>")
	'end if
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<br><td><span style='font-size: 16pt; font-family: 標楷體'> 交通選擇：</span></td>")
	Response.Write("<table border='0' width='100%' ID='Table3'>")
	Response.Write("<tr>")
	if bus=0 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　自行接送　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=1 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　1號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=2 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　6號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=3 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　8號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=4 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　9號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=5 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　11號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=6 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　15號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=7 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　16號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=8 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　20號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=9 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　26號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=10 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　29號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=11 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　30號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=12 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　33號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=13 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　36號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=14 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　39號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=15 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　42號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=16 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　60號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=17 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　9B號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	elseif bus=18 then
		Response.Write("<td bgcolor=#FBFAFA><span style='font-size: 14pt; font-family: 標楷體'>您選擇的是<u>　9C號車　</u>。</span></td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
	end if
	Response.Write("<br><td bgcolor=#FBFAFA><span style='font-size: 16pt; font-family: 標楷體'> 搭乘校車代碼：<u>　" & busstopid & "　</u>。</span></td>")
	'if bus=7 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: 標楷體'>新社國中(8人座)</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>26號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>29號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>30號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>33號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>36號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>39號車</span></td>")
	'elseif bus=8 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>新社國中(8人座)</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: 標楷體'>26號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>29號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>30號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>33號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>36號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>39號車</span></td>")
	'elseif bus=9 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>新社國中(8人座)</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>26號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: 標楷體'>29號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>30號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>33號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>36號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>39號車</span></td>")
	'elseif bus=10 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>新社國中(8人座)</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>26號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>29號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: 標楷體'>30號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>33號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>36號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>39號車</span></td>")
	'elseif bus=11 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>新社國中(8人座)</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>26號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>29號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>30號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: 標楷體'>33號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>36號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>39號車</span></td>")
	'elseif bus=12 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>新社國中(8人座)</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>26號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>29號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>30號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>33號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: 標楷體'>36號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>39號車</span></td>")
	'elseif bus=13 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>新社國中(8人座)</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>26號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>29號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>30號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>33號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>36號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: 標楷體'>39號車</span></td>")
	'else
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>新社國中(8人座)</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>26號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>29號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>30號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>33號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>36號車</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>39號車</span></td>")
	'end if
	Response.Write("</tr>")
	Response.Write("<tr>")
	'if bus=14 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: 標楷體'>39號車B線</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>42號車</span></td>")
	'elseif bus=15 then
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>39號車B線</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus'  checked><span style='font-size: 13pt; font-family: 標楷體'>42號車</span></td>")
	'else
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>39號車B線</span></td>")
	'	Response.Write("<td bgcolor=#FBFAFA><input type='radio' name='Bus' disabled><span style='font-size: 13pt; font-family: 標楷體'>42號車</span></td>")
	'end if
	Response.Write("</tr>")
	Response.Write("</table>")
	
	' --------------------------- 自習費用 total ---------------------------
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
	Response.Write("<td><p align='center'><span style='font-size: 16pt; font-family: 標楷體'> 中　　華　　民　　國　九　十　九　年　八　月　　　　日</span></td>")
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
</html>
