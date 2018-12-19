<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>end方法</title>
<style>
body{
	margin:12px;
	font-size: 12px;
}
</style>
</head>
<body>
<%
	response.Write("End方法的使用：在调用end方法之前，这些话是可以看到的。")
	response.End()
	response.Write("在end方法之后，这些话是看不到的。")
%>
这里的话，同样看不到。
</body>
</html>
