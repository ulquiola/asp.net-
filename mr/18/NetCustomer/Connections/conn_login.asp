<%
set conn=server.CreateObject("ADODB.Connection")     '�������Ӷ���
conn.open("Driver={Microsoft Access Driver (*.mdb)};PWD=111;DBQ="&_
 server.MapPath("customer.mdb"))  '�����ַ���
%>
