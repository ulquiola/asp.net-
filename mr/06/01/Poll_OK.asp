<%@language="VBScript" %>
<!--#include file="conn/conn.asp" -->
<% 
if not request.ServerVariables("REMOTE_ADDR")=request.Cookies("IPAddress") then
	response.Cookies("IPAddress")=request.ServerVariables("REMOTE_ADDR")
	response.Cookies("IPAddress").expires=date+30   '����Cookie����������
else %>
	<script language="javascript">
	alert("���Ѿ�Ͷ��Ʊ�ˣ�лл����֧�֣�");
	window.location="Poll_Browse.asp"  //ת����ʾͶƱ���ҳ��
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