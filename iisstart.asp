<%@ Language = "VBScript" %>
<% Response.Buffer = true %>

<!--
    WARNING!
    Please do not alter this file. It may be replaced if you upgrade your web server 
      If you want to use it as a template, we recommend renaming it, and modifying the new file.
    Thanks.
-->

<html>

<head>
<meta HTTP-EQUIV="Content-Type" Content="text-html; charset=big5">

<%

Dim strServername, strLocalname, strServerIP

strServername = LCase(Request.ServerVariables("SERVER_NAME"))   ' Server's name
strServerIP = LCase(Request.ServerVariables("LOCAL_ADDR"))      ' Server's IP address
strRemoteIP =  LCase(Request.ServerVariables("REMOTE_ADDR"))    ' Client's IP address

' If the querystring variable uc <> 1, and the user is browsing from the server machine, 
' go ahead and show them localstart.asp.  We don't want localstart.asp shown to outside users.

If Request("uc") <> 1 And  (strServername = "localhost" Or strServerIP = strRemoteIP) Then
  Response.Redirect "localstart.asp"
Else 

%>

<title id=titletext>建構中</title>
</head>

  <body bgcolor=white>
  <table>
  <tr>
  <td id="tableProps" width=70 valign=top align=center>
  <img id="pagerrorImg" src="pagerror.gif" width=36 height=48>  
  <td id="tablePropsWidth" width=400>
  
  <h1 id=errortype style="font:14pt/16pt 新細明體; color:#4e4e4e">
  <id id="Comment1"><!--Problem--></id><id id="errorText">建構中</id></h1>
  <id id="Comment2"><!--Probable causes:<--></id><id id="errordesc"><font style="font:9pt/12pt 新細明體; color:black">
  您試圖到達的站台目前並沒有預設頁面。可能站台正在升級及設定。
  </id>
  <br><br>
  
  <hr size=1 color="blue">
  
  <br>
  <id  id=term1>
  請稍後再重試這個站台。如果問題仍然存在，請洽詢網站管理員。
  </id>
  <p>
  
  </ul>
  <br>
  </td>
  </tr>
  </table>
  </body>
  
<%

End If 

%>

</html>














