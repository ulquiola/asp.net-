<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>request对象的使用</title>
<style>
body{
	margin:12px;
	font-size:12px;
	line-height:20px;
	word-break: break-all
}
</style>
</head>
<body>
<%
	response.Write("您机器的IP地址是："&request.ServerVariables("REMOTE_ADDR"))
	response.Write("<br/>")
	response.Write("服务器端口号："&request.ServerVariables("SERVER_PORT"))
%>
</body>
</html>
