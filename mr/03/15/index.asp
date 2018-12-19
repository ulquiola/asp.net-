<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Application的使用</title>
</head>
<body>
<%
Dim arrays(8)							'声明数组
arrays(0)="欢迎"						'为数组赋值
arrays(1)="光临"						'为数组赋值	
arrays(2)="豆豆"						'为数组赋值
arrays(3)="网站!"						'为数组赋值
application.Lock()						'应用Lock方法进行锁定
application("sun")=arrays				'为变量赋值
application.UnLock()					'解除锁定
Dim arrayys								'定义变量
application.Lock()						'应用Lock方法进行锁定
arrayys=application("sun")				'为变量赋值
application.UnLock()					'解除锁定
response.Write "<i>"&(arrayys(0)&arrayys(1)&arrayys(2)&arrayys(3)&arrayys(4)&arrayys(5)&arrayys(6)) 
'动态输出数组中指定数据信息
%>
</body>
</html>
