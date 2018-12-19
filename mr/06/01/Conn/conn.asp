<%
Set conn=Server.CreateObject("ADODB.Connection")
sql="Driver={Microsoft Access Driver (*.mdb)};DBQ= "&Server.MapPath("Database/bookshop.mdb")
conn.open(sql)
%>

