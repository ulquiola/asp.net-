<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>MapPath����</title>
<style>
body{
	margin:12px;
	font-size:12px;
}
</style>
</head>
<body>
<%
Response.Write("����������Ŀ¼������·��Ϊ��"&Server.MapPath("\")&"<BR>")
Response.Write("��ǰ����·��Ϊ��"&Server.MapPath("./")&"<BR>")
Response.Write("��Ŀ¼����·��Ϊ��"&Server.MapPath("../")&"<BR>")
Response.Write("��ǰ�ļ�����·��Ϊ��"&Server.MapPath(Request.ServerVariables("PATH_INFO"))) 
'ʹ��Request����ķ���������PATH_INFO��ӳ�䵱ǰ�ļ�������·��
%>

</body>
</html>
