<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>HTMLEncode</title>
<style>
body{
	margin:12px;
	font-size:12px;
}
</style>
</head>
<body>
<marquee width="300" height="25" direction=left>�Ҵ������������</marquee>
<p />
<%
	str = "<marquee width=&quot;300&quot; height=&quot;25&quot; direction=left>�Ҵ������������</marquee>"
	response.Write(server.HTMLEncode(str))
%>
</body>
</html>
