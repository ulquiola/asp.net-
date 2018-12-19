<%
set conn=server.CreateObject("adodb.connection")
sql="Driver={Microsoft Access Driver (*.mdb)};DBQ=" &server.MapPath("Database/db_sousuo.mdb")
conn.open(sql)
%>
