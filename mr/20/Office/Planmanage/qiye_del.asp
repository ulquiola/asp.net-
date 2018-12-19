<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="../conn/conn.asp"-->
<%
del="delete from Tab_qiye where id="&request.QueryString("id")
conn.execute(del)
%>
<script language="javascript">
{alert("企业计划信息删除成功!")}
window.location.href='qiye_planindex.asp';
</script>