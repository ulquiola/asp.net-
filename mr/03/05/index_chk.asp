<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
</head>
<body>
<%
	lgname = request.Form("lgname")
	lgpwd = request.Form("lgpwd")
	autolg = request.Form("autolg")
	
	if lgname = "" or lgpwd = "" then
		response.Write("�û������벻����Ϊ�ա���<a href="&request.ServerVariables("HTTP_REFERER")&">����</a>")
		response.End()
	end if
	if autolg = 1 then
		response.Cookies("login")("lgname") = lgname
		response.Cookies("login")("lgpwd") = lgpwd
		response.Cookies("login").Expires = now() + 60*60
	else
		response.Cookies("login").Expires = now() - 1
	end if
	response.Write("��¼�ɹ�����ӭ���٣�"&lgname)
	response.Write("<p><a href="&request.ServerVariables("HTTP_REFERER")&">����</a>����¼ҳ")
	
%>
</body>
</html>
