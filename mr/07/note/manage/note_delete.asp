<!--#include file="../common/conn.asp"-->
<%
note_id = request("note_id")		'获取留言编号
sql = "delete from tb_note where note_id in("&note_id&")"			     '删除留言信息表中的数据
sql1="delete from tb_note_answer where noan_note_id in("&note_id&")"     '删除留言回复信息表中的数据
conn.Execute(sql)
conn.Execute(sql1)
%>
<script language="javascript">alert("留言信息删除成功！");window.location.href="index.asp";</script>
<meta http-equiv="Content-Type" content="text/html;charset=gb2312">
