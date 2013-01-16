<% 
'引用本程式段請用以下語法
'  <!--#InClude File="sql_conn.asp" -->
'連接資料庫
        set conn = server.CreateObject("Adodb.connection")
        conn.Open "Provider=SQLOLEDB;Data Source=0B78;Initial Catalog=sdata; User ID=sa;Password=1234;"
        set rs = server.CreateObject("Adodb.recordset")
%>
