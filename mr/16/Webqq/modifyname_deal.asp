<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
<%
    sql="Update tb_user set state='"&request.form("state")&"' where ID="&request.cookies("UserID")&""
    Set rs=Server.CreateObject("ADODB.Command")
    rs.ActiveConnection = conn
    rs.CommandText = sql
    rs.Execute
    rs.ActiveConnection.Close
    Set rs=nothing
	Session("state")=request.form("state")
%>
<html>
<head>
<script language="javascript">
alert("祝贺您,您的昵称修改成功!");
window.opener.location.reload();//刷新父窗口中的网页
window.close();//关闭当前窗口
</script>
</head>
</html>