<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>transfer����</title>
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
'���ûỰ��ʱʱ��
Server.ScriptTimeout=60       
'����ASP�ű���ʱʱ��
Server.Transfer("include.inc")
%>
<font color="#0000FF">���ٷ��ر�ҳ��ִ��!</font>

</body>
</html>
