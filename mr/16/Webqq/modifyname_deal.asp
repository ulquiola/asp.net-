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
alert("ף����,�����ǳ��޸ĳɹ�!");
window.opener.location.reload();//ˢ�¸������е���ҳ
window.close();//�رյ�ǰ����
</script>
</head>
</html>