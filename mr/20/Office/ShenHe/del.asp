<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="../conn/conn.asp"-->
<%
del="delete from Tab_shehe where id="&request.QueryString("id")
conn.execute(del)
%>
<script language="javascript">
{alert("��Ϣɾ���ɹ�!")}
window.location.href='piguan.asp';
</script>