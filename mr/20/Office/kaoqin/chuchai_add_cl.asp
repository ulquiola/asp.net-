<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="../conn/conn.asp"-->
<%
if request.Form("name1")<>"" then
name1=request.Form("name1")
department=request.Form("department")
chuarea=request.Form("chuarea")
time1=request.Form("time1")
time2=request.Form("time2")
sql="insert into Tab_chuchai(name1,department,chuarea,time1,time2) values('"&name1&"','"&department&"','"&chuarea&"','"&time1&"','"&time2&"')"
conn.execute(sql)
%>
<script language="javascript">
alert("出差信息添加成功！");
opener.location.reload();
window.close();
</script>
<%end if%>