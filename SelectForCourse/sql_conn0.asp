<% 
'�ޥΥ��{���q�ХΥH�U�y�k
'  <!--#InClude File="sql_conn.asp" -->
'�s����Ʈw
        set conn = server.CreateObject("Adodb.connection")
        conn.Open "Provider=SQLOLEDB;Data Source=0B78;Initial Catalog=sdata; User ID=sa;Password=1234;"
        set rs = server.CreateObject("Adodb.recordset")
%>
