<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="../conn/conn.asp"-->
<%
del="delete from Tab_person where id="&request.QueryString("id")
conn.execute(del)
%>
<script language="javascript">
{alert("���˼ƻ���Ϣɾ���ɹ�!")}
window.location.href='person_index.asp';
</script>