<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>transfer方法</title>
<style>
	body{
		margin:12px;
		font-size:12px;
	}
</style>
</head>
<body>
<%
Session.Timeout=10       
'设置会话超时时间
Server.ScriptTimeout=60       
'设置ASP脚本超时时间
Server.Transfer("include.inc")
%>
<font color="#0000FF">不再返回本页面执行!</font>

</body>
</html>
