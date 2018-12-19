<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="../conn/conn.asp"-->
<%
if request.Form("name1")<>"" then
name1=request.Form("name1")
title=request.Form("title")
content=request.Form("content")
riqi=request.Form("riqi")
sql="insert into Tab_qiye(name1,title,content,riqi) values('"&name1&"','"&title&"','"&content&"','"&riqi&"')"
conn.execute(sql)
%>
<script language="javascript">
alert("企业计划发表成功!!");
opener.location.reload();
window.close();
</script>
<%end if%>