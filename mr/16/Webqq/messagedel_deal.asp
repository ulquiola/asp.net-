<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="conn/conn.asp"-->
<%
if request.QueryString("id")<>"" then
	Del="Delete from tb_talk where ID="&request.QueryString("id")
	conn.execute(Del)
%>
	<script language="javascript">
	window.location.href='message_delete.asp';
	</script>
<%end if%>