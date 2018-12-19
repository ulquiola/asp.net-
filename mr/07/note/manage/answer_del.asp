<!--#include file="../common/conn.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RecordSet")
id=request("id")
note_id=request("note_id")
sql_del = "delete from tb_note_answer where noan_id = "&id
rs.Open sql_del,conn,1,3
%>
<script language="javascript">alert("成功删除版主回复的留言信息！");window.location.href="note_read.asp?note_id=<%response.write(note_id)%>";</script>
<meta http-equiv="Content-Type" content="text/html;charset=gb2312">
