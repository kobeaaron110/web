<% Session.codepage=65001%>
<% Response.CharSet="utf-8" %>
<!--#InClude File="sql_conn.asp" -->
<!--#InClude File="bus_checkvalue.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Language" content="zh-tw">
		<meta http-equiv="Content-Type" content="text/html; charset=big5" />
		<STYLE type="text/css">
		A {text-decoration:none}
		A:hover {color: #FF0000; text-decoration: underline}
		</STYLE>
		<title>選擇夜輔暨晚自習參加科目</title>
	</head>
	<body bgcolor="#C7EDCC">
<Br><Br>
<%
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
	busval=rs("Bus")
	busid=rs("BusStopID")
	if isnull(rs("Password1")) then
		identity=""
	else
		identity=rs("Password1")
	end	if
	sql1="Select * from SystemVariable Where SysVarName = 'SelectCourseSchoolYear'"
	set rs1=conn.Execute(sql1)
	sql2="Select * from SystemVariable Where SysVarName = 'SelectCourseSemester'"
	set rs2=conn.Execute(sql2)
	
	if Request.Form("Subject1")="on" then
		Subject1=1
	else
		Subject1=0
	end if
	if Request.Form("Subject2")="on" then
		Subject2=1
	else
		Subject2=0
	end if
	if Request.Form("Subject3")="on" then
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
		sql_bus="Select * from 站名資料 inner join 輔導課校車 on BusNum=車號 Where BusCode=" & busval & " order by 站序"
		set rs_bus=conn.Execute(sql_bus)
	end if
	Bus=Request.Form("Bus")
	BusStopID=Request.Form("Select1")
	tel=Request.Form("Telephone")
	mobile=Request.Form("Mobile")
    if Request.Form("nextsubmit")="下一步" then
		if BusStopID<>"請選擇上下車站名" or Bus=0 or Bus=16 then
			if Bus>0 and Bus<16 then
				sql_bus="Select * from 站名資料 Where 代碼=" & BusStopID
				set rs_bus=conn.Execute(sql_bus)
				BusStopName=rs_bus("站名")
			else
				BusStopID=""
			end if
			if tel="" or mobile="" then
				Response.Write("<td><span style='font-size: 18pt; color:#FF0000; font-family: 標楷體'><strong>電話未填寫，請重新填寫電電話，重新選擇科目後；</strong></span><br><br></td>")
				Response.Write("<td><span style='font-size: 18pt; color:#FF0000; font-family: 標楷體'><strong>點擇校車車次後，再選擇上下車站名，再按下一步繼續！</strong></span><br><br></td>")
			else
				sql_update="Update PermRec Set Course1=" & Subject1 & ", Course2=" & Subject2 & ", Course3=" & Subject3 & ", Course4=" & Subject4
				sql_update=sql_update+", Course5=" & Subject5 & ", Course6=" & Subject6 & ", Course7=" & Subject7 & ", Bus=" & Bus & ", BusStopID='" & BusStopID
				sql_update=sql_update+"', Tel='" & tel & "', Mobile='" & mobile
				sql_update=sql_update+"' Where SNum='" & Account & "'"
				'Response.Write(BusStopID & "," & BusStopName)
				set rs_update=conn.Execute(sql_update)
				Response.Redirect "print.asp"
			end if
		else
			Response.Write("<td><span style='font-size: 18pt; color:#FF0000; font-family: 標楷體'><strong>上下車站名未選取，請重新填寫電電話，重新選擇科目後；</strong></span><br><br></td>")
			Response.Write("<td><span style='font-size: 18pt; color:#FF0000; font-family: 標楷體'><strong>點擇校車車次後，再選擇上下車站名，再按下一步繼續！</strong></span><br><br></td>")
		end if
	end if
	
	Response.Write("<table border='0' width='100%' ID='Table1'>")
	Response.Write("<form method='POST' name='Select' action='main.asp' ID='Form1'>")
	Response.Write("<tr>")
	Response.Write("<td colspan=2><p align='center'><span style='font-size: 18pt; font-family: 標楷體'> " & rs1("SysVarValue") & "學年度第" & rs2("SysVarValue") & "學期夜間輔導暨晚自習活動參加科目</span></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td colspan=2><span style='font-size: 15pt; font-family: 標楷體'> 班級：" & rs("ClassName") & "　　　　　學號：" & Account & "　　　　　　姓名：" & rs("Name") & "</span></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
    Response.Write("<td bgcolor='#008080' width='25%'><p align='center'><font color='#FFFFFF'><span style='font-size: 12pt'>請填寫聯絡電話</span></font></td>")
	Response.Write("<td><font color='#008080'><span style='font-size: 12pt'><INPUT type='text' value='" & rs("tel") & "' name='Telephone" & cnt & "' maxlength='20'></span></font></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
    Response.Write("<td bgcolor='#008080'><p align='center'><font color='#FFFFFF'><span style='font-size: 12pt'>請填寫行動電話</span></font></td>")
	Response.Write("<td><font color='#008080'><span style='font-size: 12pt'><INPUT type='text' value='" & rs("mobile") & "' name='Mobile" & cnt & "' maxlength='10'></span></font></td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	Response.Write("</tr>")
	if identity="pro" then 
		Response.Write("<td colspan=2> 參加科目：　　　　　　　　　　　　　　<a href='9802夜間輔導實施要點(高一特版).doc' target='_blank'>下載 夜間輔導暨晚自習實施要點</td>")
	elseif (gradeyear=1 or gradeyear=2) and left(classname,1)="普" then 
		Response.Write("<td colspan=2> 參加科目：　　　　　　　　　　　　　　<a href='9802夜間輔導實施要點(普通科版).doc' target='_blank'>下載 夜間輔導暨晚自習實施要點</td>")
	elseif gradeyear=1 or gradeyear=2 then 
		Response.Write("<td colspan=2> 參加科目：　　　　　　　　　　　　　　<a href='9802夜間輔導實施要點(高一二職版).doc' target='_blank'>下載 夜間輔導暨晚自習實施要點</td>")
	elseif gradeyear=3 then 
		Response.Write("<td colspan=2> 參加科目：　　　　　　　　　　　　　　<a href='9802夜間輔導實施要點(高三職版).doc' target='_blank'>下載 夜間輔導暨晚自習實施要點</td>")
	end if
	Response.Write("<table border='1' width='100%' ID='Table2'>")
	Response.Write("<tr>")
	i=i+1
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB width='50%'>")
	else 
		Response.Write("<td bgcolor=#FBFAFA width='50%'>")
	end if
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 14pt'>開課班別（請自行勾選）</span></font></td>")
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB width='50%'>")
	else 
		Response.Write("<td bgcolor=#FBFAFA width='50%'>")
	end if
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 14pt'>適合參加對象</span></font>")
	Response.Write("</td>")
	Response.Write("</tr>")
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
				Response.Write("<input type='checkbox' name='Subject1' checked> 國保晚自習班(星期一 ～ 星期五)</td>")
			else
				Response.Write("<input type='checkbox' name='Subject1'> 國保晚自習班(星期一 ～ 星期五)</td>")
			end if 
		else
			if course1=true then
				Response.Write("<input type='checkbox' name='Subject1' checked> 高一晚自習班(星期一 ～ 星期五)</td>")
			else
				Response.Write("<input type='checkbox' name='Subject1'> 高一晚自習班(星期一 ～ 星期五)</td>")
			end if 
		end if 
	elseif gradeyear=2 then
		if course1=true then
			Response.Write("<input type='checkbox' name='Subject1' checked> 高二晚自習班(星期一 ～ 星期五)</td>")
		else
			Response.Write("<input type='checkbox' name='Subject1'> 高二晚自習班(星期一 ～ 星期五)</td>")
		end if 
	else
		if course1=true then
			Response.Write("<input type='checkbox' name='Subject1' checked> 高三晚自習班(星期一 ～ 星期五)</td>")
		else
			Response.Write("<input type='checkbox' name='Subject1'> 高三晚自習班(星期一 ～ 星期五)</td>")
		end if 
	end if 
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB>")
	else 
		Response.Write("<td bgcolor=#FBFAFA>")
	end if 
	if identity="pro" then
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> 　 </span></font></td>")
	else
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> 全體學生 </span></font></td>")
	end if 
	Response.Write("</tr>")
						
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
				Response.Write("<input type='checkbox' name='Subject2' checked> 國保數學夜輔班</td>")
			else
				Response.Write("<input type='checkbox' name='Subject2'> 國保數學夜輔班</td>")
			end if 
		elseif left(classname,1)="普" then 
			if right(classname,1)="甲" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked> 普一甲數學加強班(星期一)</td>")
				else
					Response.Write("<input type='checkbox' name='Subject2'> 普一甲數學加強班(星期一)</td>")
				end if 
			elseif right(classname,1)="乙" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked> 普一乙數學加強班(星期二)</td>")
				else
					Response.Write("<input type='checkbox' name='Subject2'> 普一乙數學加強班(星期二)</td>")
				end if 
			end if 
		elseif left(classname,1)="幼" or left(classname,1)="美" then 
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked> 高一高職家事類數學 A班(星期三)</td>")
			else
				Response.Write("<input type='checkbox' name='Subject2'> 高一高職家事類數學 A班(星期三)</td>")
			end if 
		else
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked> 高一高職商業類數學 B班(星期三)</td>")
			else
				Response.Write("<input type='checkbox' name='Subject2'> 高一高職商業類數學 B班(星期三)</td>")
			end if 
		end if 
	elseif gradeyear=2 then
		if left(classname,1)="普" then 
			if right(classname,1)="甲" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked> 普二甲數學加強班(星期三)</td>")
				else
					Response.Write("<input type='checkbox' name='Subject2'> 普二甲數學加強班(星期三)</td>")
				end if 
			elseif right(classname,1)="乙" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked> 普二乙數學加強班(星期四)</td>")
				else
					Response.Write("<input type='checkbox' name='Subject2'> 普二乙數學加強班(星期四)</td>")
				end if 
			end if 
		elseif left(classname,1)="幼" or left(classname,1)="美" then 
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked> 高二高職家事類數學 A班(星期四)</td>")
			else
				Response.Write("<input type='checkbox' name='Subject2'> 高二高職家事類數學 A班(星期四)</td>")
			end if 
		else
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked> 高二高職商業類數學 B班(星期二)</td>")
			else
				Response.Write("<input type='checkbox' name='Subject2'> 高二高職商業類數學 B班(星期二)</td>")
			end if 
		end if 
	else
		if classname="高三11" or classname="高三12" or classname="高三13" or classname="高三14" or classname="高三15" or classname="高三16" then 
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked> 高三高職家事類數學 A班(星期一)</td>")
			else
				Response.Write("<input type='checkbox' name='Subject2'> 高三高職家事類數學 A班(星期一)</td>")
			end if 
		else
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked> 高三高職商業類數學 B班(星期一)</td>")
			else
				Response.Write("<input type='checkbox' name='Subject2'> 高三高職商業類數學 B班(星期一)</td>")
			end if 
		end if 
	end if 
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB>")
	else 
		Response.Write("<td bgcolor=#FBFAFA>")
	end if 
	Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> 　 </span></font></td>")
	Response.Write("</tr>")
	
	if gradeyear=1 and left(classname,1)<>"普" and identity<>"pro" then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course3=true then
			Response.Write("<input type='checkbox' name='Subject3' checked  onclick='eng_basic();'> 高一高職英文基礎班(星期一)</td>")
		else
			if course4=true then
				flag_engb="disabled"
			else
				flag_engb=""
			end if 
			Response.Write("<input type='checkbox' name='Subject3' " & flag_engb & "  onclick='eng_basic();'> 高一高職英文基礎班(星期一)</td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> 段考平均為60分(含)以下者 </span></font></td>")
		Response.Write("</tr>")
	end if 
	
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
				Response.Write("<input type='checkbox' name='Subject4' checked> 國保英文夜輔班</td>")
			else
				Response.Write("<input type='checkbox' name='Subject4'> 國保英文夜輔班</td>")
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> 　 </span></font></td>")
			Response.Write("</tr>")
		elseif left(classname,1)="普" then 
			if right(classname,1)="甲" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked> 普一甲英文加強班(星期二)</td>")
				else
					Response.Write("<input type='checkbox' name='Subject4'> 普一甲英文加強班(星期二)</td>")
				end if 
			elseif right(classname,1)="乙" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked> 普一乙英文加強班(星期五)</td>")
				else
					Response.Write("<input type='checkbox' name='Subject4'> 普一乙英文加強班(星期五)</td>")
				end if 
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> 　 </span></font></td>")
			Response.Write("</tr>")
		else
			if course4=true then
				Response.Write("<input type='checkbox' name='Subject4' checked  onclick='eng_adv();'> 高一高職英文加強班(星期一)</td>")
			else
				if course3=true then
					flag_enga="disabled"
				else
					flag_enga=""
				end if 
				Response.Write("<input type='checkbox' name='Subject4' " & flag_enga & "  onclick='eng_adv();'> 高一高職英文加強班(星期一)</td>")
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> 段考平均為60分以上者 </span></font></td>")
			Response.Write("</tr>")
		end if 
	elseif gradeyear=2 then
		if left(classname,1)="普" then 
			if right(classname,1)="甲" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked> 普二甲英文加強班(星期二)</td>")
				else
					Response.Write("<input type='checkbox' name='Subject4'> 普二甲英文加強班(星期二)</td>")
				end if 
			elseif right(classname,1)="乙" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked> 普二乙英文加強班(星期二)</td>")
				else
					Response.Write("<input type='checkbox' name='Subject4'> 普二乙英文加強班(星期二)</td>")
				end if 
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> 　 </span></font></td>")
			Response.Write("</tr>")
		end if 
	else
		if course4=true then
			Response.Write("<input type='checkbox' name='Subject4' checked> 高三高職英文加強班(星期五)</td>")
		else
			Response.Write("<input type='checkbox' name='Subject4'> 高三高職英文加強班(星期五)</td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> 為積極準備四技統測，提升英文實力之高三職學生 </span></font></td>")
		Response.Write("</tr>")
	end if 
	
	if gradeyear=1 and (left(classname,1)="廣" or left(classname,1)="畫") and identity<>"pro" then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course5=true then
			Response.Write("<input type='checkbox' name='Subject5' checked  onclick='pic_basic();'> 基礎素描班(星期四)</td>")
		else
			if course6=true then
				flag_picb="disabled"
			else
				flag_picb=""
			end if 
			Response.Write("<input type='checkbox' name='Subject5' " & flag_picb & "  onclick='pic_basic();'> 基礎素描班(星期四)</td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> 廣設、媒體科學生 </span></font></td>")
		Response.Write("</tr>")
	end if 
	
	if gradeyear=1 and (left(classname,1)="廣" or left(classname,1)="畫") and identity<>"pro" then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course6=true then
			Response.Write("<input type='checkbox' name='Subject6' checked  onclick='pic_adv();'> 進階素描班(星期四)</td>")
		else
			if course5=true then
				flag_pica="disabled"
			else
				flag_pica=""
			end if 
			Response.Write("<input type='checkbox' name='Subject6' " & flag_pica & "  onclick='pic_adv();'> 進階素描班(星期四)</td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> 廣設、媒體科學生 </span></font></td>")
		Response.Write("</tr>")
	elseif gradeyear=2 and left(classname,1)="廣" then
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course6=true then
			Response.Write("<input type='checkbox' name='Subject6' checked> 進階素描班(星期四)</td>")
		else
			Response.Write("<input type='checkbox' name='Subject6'> 進階素描班(星期四)</td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> 廣設、媒體科學生 </span></font></td>")
		Response.Write("</tr>")
	end if 
		
	if gradeyear=1 and (left(classname,1)="商" or left(classname,1)="資" or left(classname,1)="貿") and identity<>"pro" then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course7=true then
			Response.Write("<input type='checkbox' name='Subject7' checked> 高一會計加強班(星期五)</td>")
		else
			Response.Write("<input type='checkbox' name='Subject7'> 高一會計加強班(星期五)</td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> 商資貿科，欲加強會計者 </span></font></td>")
		Response.Write("</tr>")
	elseif gradeyear=2 and left(classname,1)="廣" then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course7=true then
			Response.Write("<input type='checkbox' name='Subject7' checked> 麥克筆班(星期三)</td>")
		else
			Response.Write("<input type='checkbox' name='Subject7'> 麥克筆班(星期三)</td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> 廣設、媒體科學生 </span></font></td>")
		Response.Write("</tr>")
	end if 
	
	Response.Write("</table>")
	
	Response.Write("<br><td> 交通選擇：（請選擇）　　　　　　　　　<a href='夜輔校車路線站名一覽表.xls' target='_blank'>下載 夜輔校車路線站名一覽表</a></td>")
	Response.Write("<table border='1' width='100%' ID='Table3'>")
	Response.Write("<tr>")
	if busval=0 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='0' name='Bus' checked  onclick='Bus_Stop0();'> 自行接送</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='0' name='Bus'  onclick='Bus_Stop0();'> 自行接送</td>")
	end if
	if busval=1 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='1' name='Bus' checked  onclick='Bus_Stop1();'> 1號車</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='1' name='Bus'  onclick='Bus_Stop1();'> 1號車</td>")
	end if
	if busval=2 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='2' name='Bus' checked  onclick='Bus_Stop6();'> 6號車</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='2' name='Bus'  onclick='Bus_Stop6();'> 6號車</td>")
	end if
	if busval=3 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='3' name='Bus' checked  onclick='Bus_Stop9();'> 9號車</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='3' name='Bus'  onclick='Bus_Stop9();'> 9號車</td>")
	end if
	if busval=4 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='4' name='Bus' checked  onclick='Bus_Stop15();'> 15號車</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='4' name='Bus'  onclick='Bus_Stop15();'> 15號車</td>")
	end if
	if busval=5 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='5' name='Bus' checked  onclick='Bus_Stop16();'> 16號車</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='5' name='Bus'  onclick='Bus_Stop16();'> 16號車</td>")
	end if
	if busval=6 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='6' name='Bus' checked  onclick='Bus_Stop20();'> 20號車</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='6' name='Bus'  onclick='Bus_Stop20();'> 20號車</td>")
	end if
	Response.Write("</tr>")
	Response.Write("<tr>")
	if busval=7 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='7' name='Bus' checked  onclick='Bus_Stop88();'> 新社國中(8人座)</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='7' name='Bus'  onclick='Bus_Stop88();'> 新社國中(8人座)</td>")
	end if
	if busval=8 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='8' name='Bus' checked  onclick='Bus_Stop26();'> 26號車</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='8' name='Bus'  onclick='Bus_Stop26();'> 26號車</td>")
	end if
	if busval=9 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='9' name='Bus' checked  onclick='Bus_Stop29();'> 29號車</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='9' name='Bus'  onclick='Bus_Stop29();'> 29號車</td>")
	end if
	if busval=10 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='10' name='Bus' checked  onclick='Bus_Stop30();'> 30號車</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='10' name='Bus'  onclick='Bus_Stop30();'> 30號車</td>")
	end if
	if busval=11 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='11' name='Bus' checked  onclick='Bus_Stop33();'> 33號車</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='11' name='Bus'  onclick='Bus_Stop33();'> 33號車</td>")
	end if
	if busval=12 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='12' name='Bus' checked  onclick='Bus_Stop36();'> 36號車</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='12' name='Bus'  onclick='Bus_Stop36();'> 36號車</td>")
	end if
	if busval=13 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='13' name='Bus' checked  onclick='Bus_Stop39();'> 39號車</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='13' name='Bus'  onclick='Bus_Stop39();'> 39號車</td>")
	end if
	Response.Write("</tr>")
	Response.Write("<tr>")
	if busval=14 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='14' name='Bus' checked  onclick='Bus_Stop39B();'> 39號車B線</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='14' name='Bus'  onclick='Bus_Stop39B();'> 39號車B線</td>")
	end if
	if busval=15 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='15' name='Bus' checked  onclick='Bus_Stop42();'> 42號車</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='15' name='Bus'  onclick='Bus_Stop42();'> 42號車</td>")
	end if
	if busval=16 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='16' name='Bus' checked  onclick='Bus_Stop0();'> 溪霸線</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='16' name='Bus'  onclick='Bus_Stop0();'> 溪霸線</td>")
	end if
	Response.Write("</tr>")
	Response.Write("</table>")
	
	Response.Write("<tr>")
	Response.Write("<br><td> 站名選擇：（請選擇）</td>")
	Response.Write("</tr>")
	Response.Write("<tr>")
	Response.Write("<td><br></td>")
	response.write("<SELECT id='Select1' name='Select1'>")
	response.write("<option>請選擇上下車站名</option>")
	if busval>0 then
		while not rs_bus.EOF
			if rs_bus("代碼")=busid then
				flag_select="selected"
			else
				flag_select=""
			end if
			response.write("<option value=" & rs_bus("代碼") & " " & flag_select & ">" & rs_bus("代碼") & " " & rs_bus("站名") & "</option>")
			rs_bus.MoveNext
		wend
	end if
	response.write("</SELECT>")
	Response.Write("</tr>")
	
    Response.Write("<tr>")
    Response.Write("<br><p align='center'><td width='50%' height='60'><input class='button' type='submit' value='下一步' name='nextsubmit' onclick='return checkvalue()' ID='Submit1'>")
    Response.Write("</td>")
    Response.Write("</tr>")
    
	Response.Write("</form>")
	Response.Write("</table>")
	Response.Write("</body>")
%> 
</html>
