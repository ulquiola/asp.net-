<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="../conn/conn.asp"-->
<%
if request.Form("name1")<>"" then
name1=request.Form("name1")
title=request.Form("title")
content=request.Form("content")
time1=request.Form("time1")
sql="insert into Tab_bm(name1,title,content,time1) values('"&name1&"','"&title&"','"&content&"','"&time1&"')"
conn.execute(sql)
%>
<script language="javascript">
alert("部门计划发表成功!!");
opener.location.reload();
window.close();
</script>
<%end if%>