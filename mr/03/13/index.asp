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
	response.Write("�й�����<br/>")
	response.Write("��ǰʱ��Ϊ��"&Date()&" "&Time()&"<br>")
	response.Write("��ǰ�ļ۸�Ϊ��"&FormatCurrency(59.3)&"<br/>")
	response.Write("��ǰLCID��ֵ�ǣ�"&session.LCID)
%>
</td><td>
<%
	session.LCID = 1033
	response.Write("��������<br/>")
	response.Write("��ǰʱ��Ϊ��"&Date()&" "&Time()&"<br>")
	response.Write("��ǰ�ļ۸�Ϊ��"&FormatCurrency(59.3)&"<br/>")
	response.Write("��ǰLCID��ֵ�ǣ�"&session.LCID)
	
%>
</td></tr>
</table>
</body>
</html>
