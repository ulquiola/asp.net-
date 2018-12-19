<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="conn/conn.asp"-->
<%  set rs=server.CreateObject("adodb.recordset")
	Del="Delete from tb_talk"
	conn.execute(Del)
%>
	<script language="javascript">
	window.location.href='message_delete.asp';
	</script>
