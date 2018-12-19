<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="../conn/conn.asp"-->
<%
if request.Form("name1")<>"" then
name1=request.Form("name1")
department=request.Form("department")
content=request.Form("content")
time1=request.Form("time1")
time2=request.Form("time2")
sql="insert into Tab_waichu(name1,department,content,time1,time2) values('"&name1&"','"&department&"','"&content&"','"&time1&"','"&time2&"')"
conn.execute(sql)
%>
<script language="javascript">
alert("外出信息添加成功！");
opener.location.reload();
window.close();
</script>
<%end if%>