<%
Set conn=Server.CreateObject("ADODB.Connection")
conn.connectionstring="Driver={Sql Server};Server=(local);UID=sa;database=db_leaveword"
conn.open
%>