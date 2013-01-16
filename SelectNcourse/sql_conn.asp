<% 
'引用本程式段請用以下語法
'  <!--#InClude File="sql_conn.asp" -->
'連接資料庫
        set conn = server.CreateObject("Adodb.connection")
        'conn.Open "Provider=SQLOLEDB;Data Source=203.68.204.6;Initial Catalog=ecampusd; User ID=sa;Password=yang8742;"
		conn.Open "Provider=SQLOLEDB;Data Source=aaron;Initial Catalog=ecampusd; User ID=sa;Password=bc55;"
        set rs = server.CreateObject("Adodb.recordset")
%>
