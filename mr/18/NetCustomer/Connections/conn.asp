<%
set conn=server.CreateObject("ADODB.Connection")
conn.open("Driver={Microsoft Access Driver (*.mdb)};PWD=111;DBQ="& server.MapPath("../customer.mdb"))
%>
