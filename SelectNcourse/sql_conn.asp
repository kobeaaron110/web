<% 
'�ޥΥ��{���q�ХΥH�U�y�k
'  <!--#InClude File="sql_conn.asp" -->
'�s����Ʈw
        set conn = server.CreateObject("Adodb.connection")
        'conn.Open "Provider=SQLOLEDB;Data Source=203.68.204.6;Initial Catalog=ecampusd; User ID=sa;Password=yang8742;"
		conn.Open "Provider=SQLOLEDB;Data Source=aaron;Initial Catalog=ecampusd; User ID=sa;Password=bc55;"
        set rs = server.CreateObject("Adodb.recordset")
%>
