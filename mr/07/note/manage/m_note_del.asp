<!--#include file="../common/conn.asp"-->
<%
note_id = request("note_id")		'��ȡ���
sql = "delete from tb_note where note_id in("&note_id&")"
conn.Execute(sql)
%>
<script language="javascript">alert("������Ϣɾ���ɹ���");window.location.href="manage_note.asp";</script>
<meta http-equiv="Content-Type" content="text/html;charset=gb2312">
 	
	
	