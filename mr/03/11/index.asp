<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Contents集合</title>
<style type="text/css">
	body{
		margin:12px;
		font-size:12px;
	}
</style>
</head>
<body>
<%
session("bccd")="www.mrbccd.com"								'为Session变量赋值
session("bcty")="www.bcty365.com"								'为Session变量赋值
session("bc110")="www.bc110.com"								'为Session变量赋值
	For each key in Session.Contents							'进行循环显示
		Response.Write "Contents数据集合的成员:"&key&"<BR>"		'输出Contents数据集合的成员
	Next
%>
</body>
</html>
