<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>ʹ��ScriptTimeout����</title>
</head>
<body>
<%
response.Write "��ҳ�����е��ʱ��Ϊ"&server.ScriptTimeout&"�롣<br>"
'���ҳ�����е��ʱ���Ƕ�����
%>
<%
Dim sun,ssml				'�������
Randomize    			'��ʼ�������������
Do Until ssml = vbno			'�����񡱰�ť������ʱֹͣѭ��
   sun = Int((6 * Rnd) + 1)	'����1��6֮��������
Loop
%>
</body>
</html>
