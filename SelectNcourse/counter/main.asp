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
		<title>��ܩ]���[�ߦ۲߰ѥ[���</title>
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
		sql_bus="Select * from ���W��� inner join ���ɽҮը� on BusNum=���� Where BusCode=" & busval & " order by ����"
		set rs_bus=conn.Execute(sql_bus)
	end if
	Bus=Request.Form("Bus")
	BusStopID=Request.Form("Select1")
	tel=Request.Form("Telephone")
	mobile=Request.Form("Mobile")
    if Request.Form("nextsubmit")="�U�@�B" then
		if BusStopID<>"�п�ܤW�U�����W" or Bus=0 or Bus=16 then
			if Bus>0 and Bus<16 then
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
				sql_update="Update PermRec Set Course1=" & Subject1 & ", Course2=" & Subject2 & ", Course3=" & Subject3 & ", Course4=" & Subject4
				sql_update=sql_update+", Course5=" & Subject5 & ", Course6=" & Subject6 & ", Course7=" & Subject7 & ", Bus=" & Bus & ", BusStopID='" & BusStopID
				sql_update=sql_update+"', Tel='" & tel & "', Mobile='" & mobile
				sql_update=sql_update+"' Where SNum='" & Account & "'"
				'Response.Write(BusStopID & "," & BusStopName)
				set rs_update=conn.Execute(sql_update)
				Response.Redirect "print.asp"
			end if
		else
			Response.Write("<td><span style='font-size: 18pt; color:#FF0000; font-family: �з���'><strong>�W�U�����W������A�Э��s��g�q�q�ܡA���s��ܬ�ث�F</strong></span><br><br></td>")
			Response.Write("<td><span style='font-size: 18pt; color:#FF0000; font-family: �з���'><strong>�I�ܮը�������A�A��ܤW�U�����W�A�A���U�@�B�~��I</strong></span><br><br></td>")
		end if
	end if
	
	Response.Write("<table border='0' width='100%' ID='Table1'>")
	Response.Write("<form method='POST' name='Select' action='main.asp' ID='Form1'>")
	Response.Write("<tr>")
	Response.Write("<td colspan=2><p align='center'><span style='font-size: 18pt; font-family: �з���'> " & rs1("SysVarValue") & "�Ǧ~�ײ�" & rs2("SysVarValue") & "�Ǵ��]�����ɺ[�ߦ۲߬��ʰѥ[���</span></td>")
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
	if identity="pro" then 
		Response.Write("<td colspan=2> �ѥ[��ءG�@�@�@�@�@�@�@�@�@�@�@�@�@�@<a href='9802�]�����ɹ�I�n�I(���@�S��).doc' target='_blank'>�U�� �]�����ɺ[�ߦ۲߹�I�n�I</td>")
	elseif (gradeyear=1 or gradeyear=2) and left(classname,1)="��" then 
		Response.Write("<td colspan=2> �ѥ[��ءG�@�@�@�@�@�@�@�@�@�@�@�@�@�@<a href='9802�]�����ɹ�I�n�I(���q�쪩).doc' target='_blank'>�U�� �]�����ɺ[�ߦ۲߹�I�n�I</td>")
	elseif gradeyear=1 or gradeyear=2 then 
		Response.Write("<td colspan=2> �ѥ[��ءG�@�@�@�@�@�@�@�@�@�@�@�@�@�@<a href='9802�]�����ɹ�I�n�I(���@�G¾��).doc' target='_blank'>�U�� �]�����ɺ[�ߦ۲߹�I�n�I</td>")
	elseif gradeyear=3 then 
		Response.Write("<td colspan=2> �ѥ[��ءG�@�@�@�@�@�@�@�@�@�@�@�@�@�@<a href='9802�]�����ɹ�I�n�I(���T¾��).doc' target='_blank'>�U�� �]�����ɺ[�ߦ۲߹�I�n�I</td>")
	end if
	Response.Write("<table border='1' width='100%' ID='Table2'>")
	Response.Write("<tr>")
	i=i+1
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB width='50%'>")
	else 
		Response.Write("<td bgcolor=#FBFAFA width='50%'>")
	end if
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 14pt'>�}�үZ�O�]�Цۦ�Ŀ�^</span></font></td>")
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB width='50%'>")
	else 
		Response.Write("<td bgcolor=#FBFAFA width='50%'>")
	end if
	Response.Write("<p align='center'><font color='#000000'><span style='font-size: 14pt'>�A�X�ѥ[��H</span></font>")
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
				Response.Write("<input type='checkbox' name='Subject1' checked> ��O�ߦ۲߯Z(�P���@ �� �P����)</td>")
			else
				Response.Write("<input type='checkbox' name='Subject1'> ��O�ߦ۲߯Z(�P���@ �� �P����)</td>")
			end if 
		else
			if course1=true then
				Response.Write("<input type='checkbox' name='Subject1' checked> ���@�ߦ۲߯Z(�P���@ �� �P����)</td>")
			else
				Response.Write("<input type='checkbox' name='Subject1'> ���@�ߦ۲߯Z(�P���@ �� �P����)</td>")
			end if 
		end if 
	elseif gradeyear=2 then
		if course1=true then
			Response.Write("<input type='checkbox' name='Subject1' checked> ���G�ߦ۲߯Z(�P���@ �� �P����)</td>")
		else
			Response.Write("<input type='checkbox' name='Subject1'> ���G�ߦ۲߯Z(�P���@ �� �P����)</td>")
		end if 
	else
		if course1=true then
			Response.Write("<input type='checkbox' name='Subject1' checked> ���T�ߦ۲߯Z(�P���@ �� �P����)</td>")
		else
			Response.Write("<input type='checkbox' name='Subject1'> ���T�ߦ۲߯Z(�P���@ �� �P����)</td>")
		end if 
	end if 
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB>")
	else 
		Response.Write("<td bgcolor=#FBFAFA>")
	end if 
	if identity="pro" then
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> �@ </span></font></td>")
	else
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> ����ǥ� </span></font></td>")
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
				Response.Write("<input type='checkbox' name='Subject2' checked> ��O�ƾǩ]���Z</td>")
			else
				Response.Write("<input type='checkbox' name='Subject2'> ��O�ƾǩ]���Z</td>")
			end if 
		elseif left(classname,1)="��" then 
			if right(classname,1)="��" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked> ���@�Ҽƾǥ[�j�Z(�P���@)</td>")
				else
					Response.Write("<input type='checkbox' name='Subject2'> ���@�Ҽƾǥ[�j�Z(�P���@)</td>")
				end if 
			elseif right(classname,1)="�A" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked> ���@�A�ƾǥ[�j�Z(�P���G)</td>")
				else
					Response.Write("<input type='checkbox' name='Subject2'> ���@�A�ƾǥ[�j�Z(�P���G)</td>")
				end if 
			end if 
		elseif left(classname,1)="��" or left(classname,1)="��" then 
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked> ���@��¾�a�����ƾ� A�Z(�P���T)</td>")
			else
				Response.Write("<input type='checkbox' name='Subject2'> ���@��¾�a�����ƾ� A�Z(�P���T)</td>")
			end if 
		else
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked> ���@��¾�ӷ~���ƾ� B�Z(�P���T)</td>")
			else
				Response.Write("<input type='checkbox' name='Subject2'> ���@��¾�ӷ~���ƾ� B�Z(�P���T)</td>")
			end if 
		end if 
	elseif gradeyear=2 then
		if left(classname,1)="��" then 
			if right(classname,1)="��" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked> ���G�Ҽƾǥ[�j�Z(�P���T)</td>")
				else
					Response.Write("<input type='checkbox' name='Subject2'> ���G�Ҽƾǥ[�j�Z(�P���T)</td>")
				end if 
			elseif right(classname,1)="�A" then 
				if course2=true then
					Response.Write("<input type='checkbox' name='Subject2' checked> ���G�A�ƾǥ[�j�Z(�P���|)</td>")
				else
					Response.Write("<input type='checkbox' name='Subject2'> ���G�A�ƾǥ[�j�Z(�P���|)</td>")
				end if 
			end if 
		elseif left(classname,1)="��" or left(classname,1)="��" then 
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked> ���G��¾�a�����ƾ� A�Z(�P���|)</td>")
			else
				Response.Write("<input type='checkbox' name='Subject2'> ���G��¾�a�����ƾ� A�Z(�P���|)</td>")
			end if 
		else
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked> ���G��¾�ӷ~���ƾ� B�Z(�P���G)</td>")
			else
				Response.Write("<input type='checkbox' name='Subject2'> ���G��¾�ӷ~���ƾ� B�Z(�P���G)</td>")
			end if 
		end if 
	else
		if classname="���T11" or classname="���T12" or classname="���T13" or classname="���T14" or classname="���T15" or classname="���T16" then 
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked> ���T��¾�a�����ƾ� A�Z(�P���@)</td>")
			else
				Response.Write("<input type='checkbox' name='Subject2'> ���T��¾�a�����ƾ� A�Z(�P���@)</td>")
			end if 
		else
			if course2=true then
				Response.Write("<input type='checkbox' name='Subject2' checked> ���T��¾�ӷ~���ƾ� B�Z(�P���@)</td>")
			else
				Response.Write("<input type='checkbox' name='Subject2'> ���T��¾�ӷ~���ƾ� B�Z(�P���@)</td>")
			end if 
		end if 
	end if 
	if (i mod 2)=1 then 
		Response.Write("<td bgcolor=#EBEBEB>")
	else 
		Response.Write("<td bgcolor=#FBFAFA>")
	end if 
	Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> �@ </span></font></td>")
	Response.Write("</tr>")
	
	if gradeyear=1 and left(classname,1)<>"��" and identity<>"pro" then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course3=true then
			Response.Write("<input type='checkbox' name='Subject3' checked  onclick='eng_basic();'> ���@��¾�^���¦�Z(�P���@)</td>")
		else
			if course4=true then
				flag_engb="disabled"
			else
				flag_engb=""
			end if 
			Response.Write("<input type='checkbox' name='Subject3' " & flag_engb & "  onclick='eng_basic();'> ���@��¾�^���¦�Z(�P���@)</td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> �q�ҥ�����60��(�t)�H�U�� </span></font></td>")
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
				Response.Write("<input type='checkbox' name='Subject4' checked> ��O�^��]���Z</td>")
			else
				Response.Write("<input type='checkbox' name='Subject4'> ��O�^��]���Z</td>")
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> �@ </span></font></td>")
			Response.Write("</tr>")
		elseif left(classname,1)="��" then 
			if right(classname,1)="��" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked> ���@�ҭ^��[�j�Z(�P���G)</td>")
				else
					Response.Write("<input type='checkbox' name='Subject4'> ���@�ҭ^��[�j�Z(�P���G)</td>")
				end if 
			elseif right(classname,1)="�A" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked> ���@�A�^��[�j�Z(�P����)</td>")
				else
					Response.Write("<input type='checkbox' name='Subject4'> ���@�A�^��[�j�Z(�P����)</td>")
				end if 
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> �@ </span></font></td>")
			Response.Write("</tr>")
		else
			if course4=true then
				Response.Write("<input type='checkbox' name='Subject4' checked  onclick='eng_adv();'> ���@��¾�^��[�j�Z(�P���@)</td>")
			else
				if course3=true then
					flag_enga="disabled"
				else
					flag_enga=""
				end if 
				Response.Write("<input type='checkbox' name='Subject4' " & flag_enga & "  onclick='eng_adv();'> ���@��¾�^��[�j�Z(�P���@)</td>")
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> �q�ҥ�����60���H�W�� </span></font></td>")
			Response.Write("</tr>")
		end if 
	elseif gradeyear=2 then
		if left(classname,1)="��" then 
			if right(classname,1)="��" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked> ���G�ҭ^��[�j�Z(�P���G)</td>")
				else
					Response.Write("<input type='checkbox' name='Subject4'> ���G�ҭ^��[�j�Z(�P���G)</td>")
				end if 
			elseif right(classname,1)="�A" then 
				if course4=true then
					Response.Write("<input type='checkbox' name='Subject4' checked> ���G�A�^��[�j�Z(�P���G)</td>")
				else
					Response.Write("<input type='checkbox' name='Subject4'> ���G�A�^��[�j�Z(�P���G)</td>")
				end if 
			end if 
			if (i mod 2)=1 then 
				Response.Write("<td bgcolor=#EBEBEB>")
			else 
				Response.Write("<td bgcolor=#FBFAFA>")
			end if 
			Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> �@ </span></font></td>")
			Response.Write("</tr>")
		end if 
	else
		if course4=true then
			Response.Write("<input type='checkbox' name='Subject4' checked> ���T��¾�^��[�j�Z(�P����)</td>")
		else
			Response.Write("<input type='checkbox' name='Subject4'> ���T��¾�^��[�j�Z(�P����)</td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> ���n���ǳƥ|�޲δ��A���ɭ^���O�����T¾�ǥ� </span></font></td>")
		Response.Write("</tr>")
	end if 
	
	if gradeyear=1 and (left(classname,1)="�s" or left(classname,1)="�e") and identity<>"pro" then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course5=true then
			Response.Write("<input type='checkbox' name='Subject5' checked  onclick='pic_basic();'> ��¦���y�Z(�P���|)</td>")
		else
			if course6=true then
				flag_picb="disabled"
			else
				flag_picb=""
			end if 
			Response.Write("<input type='checkbox' name='Subject5' " & flag_picb & "  onclick='pic_basic();'> ��¦���y�Z(�P���|)</td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> �s�]�B�C���ǥ� </span></font></td>")
		Response.Write("</tr>")
	end if 
	
	if gradeyear=1 and (left(classname,1)="�s" or left(classname,1)="�e") and identity<>"pro" then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course6=true then
			Response.Write("<input type='checkbox' name='Subject6' checked  onclick='pic_adv();'> �i�����y�Z(�P���|)</td>")
		else
			if course5=true then
				flag_pica="disabled"
			else
				flag_pica=""
			end if 
			Response.Write("<input type='checkbox' name='Subject6' " & flag_pica & "  onclick='pic_adv();'> �i�����y�Z(�P���|)</td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> �s�]�B�C���ǥ� </span></font></td>")
		Response.Write("</tr>")
	elseif gradeyear=2 and left(classname,1)="�s" then
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course6=true then
			Response.Write("<input type='checkbox' name='Subject6' checked> �i�����y�Z(�P���|)</td>")
		else
			Response.Write("<input type='checkbox' name='Subject6'> �i�����y�Z(�P���|)</td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> �s�]�B�C���ǥ� </span></font></td>")
		Response.Write("</tr>")
	end if 
		
	if gradeyear=1 and (left(classname,1)="��" or left(classname,1)="��" or left(classname,1)="�T") and identity<>"pro" then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course7=true then
			Response.Write("<input type='checkbox' name='Subject7' checked> ���@�|�p�[�j�Z(�P����)</td>")
		else
			Response.Write("<input type='checkbox' name='Subject7'> ���@�|�p�[�j�Z(�P����)</td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> �Ӹ�T��A���[�j�|�p�� </span></font></td>")
		Response.Write("</tr>")
	elseif gradeyear=2 and left(classname,1)="�s" then 
		i=i+1
		Response.Write("<tr>")
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		if course7=true then
			Response.Write("<input type='checkbox' name='Subject7' checked> ���J���Z(�P���T)</td>")
		else
			Response.Write("<input type='checkbox' name='Subject7'> ���J���Z(�P���T)</td>")
		end if 
		if (i mod 2)=1 then 
			Response.Write("<td bgcolor=#EBEBEB>")
		else 
			Response.Write("<td bgcolor=#FBFAFA>")
		end if 
		Response.Write("<p align='left'><font color='#000000'><span style='font-size: 12pt'> �s�]�B�C���ǥ� </span></font></td>")
		Response.Write("</tr>")
	end if 
	
	Response.Write("</table>")
	
	Response.Write("<br><td> ��q��ܡG�]�п�ܡ^�@�@�@�@�@�@�@�@�@<a href='�]���ը����u���W�@����.xls' target='_blank'>�U�� �]���ը����u���W�@����</a></td>")
	Response.Write("<table border='1' width='100%' ID='Table3'>")
	Response.Write("<tr>")
	if busval=0 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='0' name='Bus' checked  onclick='Bus_Stop0();'> �ۦ汵�e</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='0' name='Bus'  onclick='Bus_Stop0();'> �ۦ汵�e</td>")
	end if
	if busval=1 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='1' name='Bus' checked  onclick='Bus_Stop1();'> 1����</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='1' name='Bus'  onclick='Bus_Stop1();'> 1����</td>")
	end if
	if busval=2 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='2' name='Bus' checked  onclick='Bus_Stop6();'> 6����</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='2' name='Bus'  onclick='Bus_Stop6();'> 6����</td>")
	end if
	if busval=3 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='3' name='Bus' checked  onclick='Bus_Stop9();'> 9����</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='3' name='Bus'  onclick='Bus_Stop9();'> 9����</td>")
	end if
	if busval=4 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='4' name='Bus' checked  onclick='Bus_Stop15();'> 15����</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='4' name='Bus'  onclick='Bus_Stop15();'> 15����</td>")
	end if
	if busval=5 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='5' name='Bus' checked  onclick='Bus_Stop16();'> 16����</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='5' name='Bus'  onclick='Bus_Stop16();'> 16����</td>")
	end if
	if busval=6 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='6' name='Bus' checked  onclick='Bus_Stop20();'> 20����</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='6' name='Bus'  onclick='Bus_Stop20();'> 20����</td>")
	end if
	Response.Write("</tr>")
	Response.Write("<tr>")
	if busval=7 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='7' name='Bus' checked  onclick='Bus_Stop88();'> �s���ꤤ(8�H�y)</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='7' name='Bus'  onclick='Bus_Stop88();'> �s���ꤤ(8�H�y)</td>")
	end if
	if busval=8 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='8' name='Bus' checked  onclick='Bus_Stop26();'> 26����</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='8' name='Bus'  onclick='Bus_Stop26();'> 26����</td>")
	end if
	if busval=9 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='9' name='Bus' checked  onclick='Bus_Stop29();'> 29����</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='9' name='Bus'  onclick='Bus_Stop29();'> 29����</td>")
	end if
	if busval=10 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='10' name='Bus' checked  onclick='Bus_Stop30();'> 30����</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='10' name='Bus'  onclick='Bus_Stop30();'> 30����</td>")
	end if
	if busval=11 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='11' name='Bus' checked  onclick='Bus_Stop33();'> 33����</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='11' name='Bus'  onclick='Bus_Stop33();'> 33����</td>")
	end if
	if busval=12 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='12' name='Bus' checked  onclick='Bus_Stop36();'> 36����</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='12' name='Bus'  onclick='Bus_Stop36();'> 36����</td>")
	end if
	if busval=13 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='13' name='Bus' checked  onclick='Bus_Stop39();'> 39����</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='13' name='Bus'  onclick='Bus_Stop39();'> 39����</td>")
	end if
	Response.Write("</tr>")
	Response.Write("<tr>")
	if busval=14 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='14' name='Bus' checked  onclick='Bus_Stop39B();'> 39����B�u</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='14' name='Bus'  onclick='Bus_Stop39B();'> 39����B�u</td>")
	end if
	if busval=15 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='15' name='Bus' checked  onclick='Bus_Stop42();'> 42����</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='15' name='Bus'  onclick='Bus_Stop42();'> 42����</td>")
	end if
	if busval=16 then
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='16' name='Bus' checked  onclick='Bus_Stop0();'> ���Q�u</td>")
	else
		Response.Write("<td bgcolor=#FBFAFA><input type='radio' value='16' name='Bus'  onclick='Bus_Stop0();'> ���Q�u</td>")
	end if
	Response.Write("</tr>")
	Response.Write("</table>")
	
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
    Response.Write("<br><p align='center'><td width='50%' height='60'><input class='button' type='submit' value='�U�@�B' name='nextsubmit' onclick='return checkvalue()' ID='Submit1'>")
    Response.Write("</td>")
    Response.Write("</tr>")
    
	Response.Write("</form>")
	Response.Write("</table>")
	Response.Write("</body>")
%> 
</html>
