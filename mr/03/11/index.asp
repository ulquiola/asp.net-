<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Contents����</title>
<style type="text/css">
	body{
		margin:12px;
		font-size:12px;
	}
</style>
</head>
<body>
<%
session("bccd")="www.mrbccd.com"								'ΪSession������ֵ
session("bcty")="www.bcty365.com"								'ΪSession������ֵ
session("bc110")="www.bc110.com"								'ΪSession������ֵ
	For each key in Session.Contents							'����ѭ����ʾ
		Response.Write "Contents���ݼ��ϵĳ�Ա:"&key&"<BR>"		'���Contents���ݼ��ϵĳ�Ա
	Next
%>
</body>
</html>
