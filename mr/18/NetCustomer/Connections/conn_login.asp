<%
set conn=server.CreateObject("ADODB.Connection")     '创建连接对象
conn.open("Driver={Microsoft Access Driver (*.mdb)};PWD=111;DBQ="&_
 server.MapPath("customer.mdb"))  '连接字符串
%>
