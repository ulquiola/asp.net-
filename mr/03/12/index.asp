<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
</head>
<body>
<%
	Dim key										'定义变量
	For each key in Session.StaticObjects		'应用For each循环输出StaticObjects数据集合遍历的对象
		Response.Write("集合的成员分别是："&key&"<BR>")	'输出数据信息
	Next
%>
</body>
</html>
