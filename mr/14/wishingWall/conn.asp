<%
set conn=Server.CreateObject("ADODB.Connection")
sql="provider=microsoft.jet.oledb.4.0;data source="& Server.mappath("database.mdb")
conn.open(sql)
%>