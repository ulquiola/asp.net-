<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open "Driver={Sql Server};Server=(local);UID=sa;PWD=;Database=db_pet"
%>