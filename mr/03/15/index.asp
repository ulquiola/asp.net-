<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Application��ʹ��</title>
</head>
<body>
<%
Dim arrays(8)							'��������
arrays(0)="��ӭ"						'Ϊ���鸳ֵ
arrays(1)="����"						'Ϊ���鸳ֵ	
arrays(2)="����"						'Ϊ���鸳ֵ
arrays(3)="��վ!"						'Ϊ���鸳ֵ
application.Lock()						'Ӧ��Lock������������
application("sun")=arrays				'Ϊ������ֵ
application.UnLock()					'�������
Dim arrayys								'�������
application.Lock()						'Ӧ��Lock������������
arrayys=application("sun")				'Ϊ������ֵ
application.UnLock()					'�������
response.Write "<i>"&(arrayys(0)&arrayys(1)&arrayys(2)&arrayys(3)&arrayys(4)&arrayys(5)&arrayys(6)) 
'��̬���������ָ��������Ϣ
%>
</body>
</html>
