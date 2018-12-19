<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="../conn/conn.asp"-->
<%
del="delete from Tab_tongxun where id="&request.QueryString("id")
conn.execute(del)
%>
<script language="javascript">
alert("通讯组删除成功!!");
window.location.href='tongxun_index.asp';
</script>