<!--#include file="common/conn.asp"-->
<%
	set rs=server.CreateObject("adodb.recordset")
	user_name = request.form("user_name")	
	title = request.form("title")
	content =request.form("content")
	h = request.form("head")&".gif"
	note_flag=request.form("checkbox")
	if note_flag<>1 then 
		note_flag=0
	end if
	datetime=now()
	sql = "insert into tb_note(note_user,note_title,note_content,note_time,note_user_pic,note_flag) values('"&user_name&"','"&title&"','"&content&"','"&datetime&"','"&h&"','"&note_flag&"')"
	conn.execute(sql)
%>	
<script language="javascript">alert("¡Ù—‘≥…π¶£°");window.location.href='index.asp';</script>
