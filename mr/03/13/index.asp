<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>LCID</title>
<style>
	body{
	margin:12px;
	font-size:12px;
	}
</style>
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0" width="450px">
<tr>
<td>
<%
	session.LCID = 2052
	response.Write("中国区域<br/>")
	response.Write("当前时间为："&Date()&" "&Time()&"<br>")
	response.Write("当前的价格为："&FormatCurrency(59.3)&"<br/>")
	response.Write("当前LCID的值是："&session.LCID)
%>
</td><td>
<%
	session.LCID = 1033
	response.Write("美国区域<br/>")
	response.Write("当前时间为："&Date()&" "&Time()&"<br>")
	response.Write("当前的价格为："&FormatCurrency(59.3)&"<br/>")
	response.Write("当前LCID的值是："&session.LCID)
	
%>
</td></tr>
</table>
</body>
</html>
