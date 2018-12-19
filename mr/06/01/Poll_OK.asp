<%@language="VBScript" %>
<!--#include file="conn/conn.asp" -->
<% 
if not request.ServerVariables("REMOTE_ADDR")=request.Cookies("IPAddress") then
	response.Cookies("IPAddress")=request.ServerVariables("REMOTE_ADDR")
	response.Cookies("IPAddress").expires=date+30   '设置Cookie的生命周期
else %>
	<script language="javascript">
	alert("您已经投过票了，谢谢您的支持！");
	window.location="Poll_Browse.asp"  //转到显示投票结果页面
	</script>
	<% response.End()
end if

if(Request.form("myif") <> "") then PID = Request.form("myif")
sql = "UPDATE DB_Poll  SET poll = poll+1  WHERE id=" + Replace(PID, "'", "''") + " "
conn.Execute sql
%>
<script language="javascript">
window.location="Poll_Browse.asp"
</script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
</html>