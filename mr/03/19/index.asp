<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>URLEncode方法</title>
<style>
body{
	margin:12px;
	font-size:12px;
}
</style>
</head>
<body>
<div style="height:25px; text-align:center">输入要转换的URL</div>
<div align="center">
<form method="post" action="index.asp">
	<input name="url" type="text" style=" width: 300px;" />
	<input type="submit" valu="转换" />
</form>
</div>
<div>
<%
	url = request.Form("url")
	if url <> "" then
		response.Write("转换前："&url&"<br />")
		response.Write("转换后："&server.URLEncode(url)&"<br>")
	end if
%>
</div>
</body>
</html>
