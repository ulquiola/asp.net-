<% response.Buffer=true %>
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<style>
body{
	margin:12px;
	font-size:12px;
}
</style>
</head>
<body>
<% 
	response.Write("这部分内容将被保存到缓存中，如果先执行Flush方法，则能够被输出；如果先执行Clear方法，这句话将不会显示")
	response.Clear()
	response.Write("<p />无论之前使用过什么方法，这句话都会显示出来。")
%>
</body>
</html>
