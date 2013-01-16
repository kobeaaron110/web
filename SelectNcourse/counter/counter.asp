<%
 Set con=Server.CreateObject("ADODB.Connection")
 con.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("counter/counter.mdb")

 SQL = "SELECT * FROM counter"
 SET RS = con.Execute(SQL)
 RS.MoveFirst
 count = RS("counts")
 If IsEmpty(Session("Conn")) Then
	Application.Lock
	Application("counter") = RS("counts") + 1
	Application.UnLock
	count = Application("counter")

	SQL1 = "UPDATE counter SET counts = " & count
	SET RS = con.Execute(SQL1)
 End If
 con.CLOSE
Session("Conn") = True
%>
