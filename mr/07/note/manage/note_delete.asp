<!--#include file="../common/conn.asp"-->
<%
note_id = request("note_id")		'��ȡ���Ա��
sql = "delete from tb_note where note_id in("&note_id&")"			     'ɾ��������Ϣ���е�����
sql1="delete from tb_note_answer where noan_note_id in("&note_id&")"     'ɾ�����Իظ���Ϣ���е�����
conn.Execute(sql)
conn.Execute(sql1)
%>
<script language="javascript">alert("������Ϣɾ���ɹ���");window.location.href="index.asp";</script>
<meta http-equiv="Content-Type" content="text/html;charset=gb2312">
