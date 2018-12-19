<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
<%
		
		sql_1 = "delete   FROM tb_talk where (fuserid = "&request.QueryString("UserID")&" and ToUserID="&request.Cookies("UserID")&") or (fuserid="&request.Cookies("UserID")&" and ToUserID=" &request.QueryString("UserID")&") "
		 Set rs_2 = Server.CreateObject("ADODB.Command")
    rs_2.ActiveConnection = conn
    rs_2.CommandText = sql_1
    rs_2.Execute
    rs_2.ActiveConnection.Close
    Set rs_2 = nothing
%>
<script>
alert("聊天记录清除成功!")
window.parent.location.reload()
</script>
