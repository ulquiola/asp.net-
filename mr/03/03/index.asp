<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>TotalBytes属性</title>
</head>
<body>
<form method="post" action="index.asp" enctype="multipart/form-data">
	文件名称：<input type="text" name="fname" /><br />
	上传文件：<input type="file" name="cont" />
	<input type="submit" value="上传" />
</form>
<%
	if request.TotalBytes <> 0 then
		response.Write("请求页的大小："&request.TotalBytes)
	end if
%>
</body>
</html>
