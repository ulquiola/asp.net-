<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
<%
	sql  = "Update tb_user set Password1='"&request.form("Password")&"' where ID="&request.cookies("UserID")&""
    Set rs1 = Server.CreateObject("ADODB.Command")
    rs1.ActiveConnection = conn
    rs1.CommandText = sql
	rs1.Execute
    rs1.ActiveConnection.Close
    Set rs1 = nothing
%>
<html>
<head>
<script>
alert("×£ºØÄú,ÃÜÂëÐÞ¸Ä³É¹¦!");
window.close();
</script>
</head>
<body>
</body>
</html>