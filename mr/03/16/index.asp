<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>使用ScriptTimeout属性</title>
</head>
<body>
<%
response.Write "该页面运行的最长时间为"&server.ScriptTimeout&"秒。<br>"
'输出页面运行的最长时间是多少秒
%>
<%
Dim sun,ssml				'定义变量
Randomize    			'初始化随机数生成器
Do Until ssml = vbno			'当“否”按钮被单击时停止循环
   sun = Int((6 * Rnd) + 1)	'产生1到6之间的随机数
Loop
%>
</body>
</html>
