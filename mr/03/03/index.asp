<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>TotalBytes����</title>
</head>
<body>
<form method="post" action="index.asp" enctype="multipart/form-data">
	�ļ����ƣ�<input type="text" name="fname" /><br />
	�ϴ��ļ���<input type="file" name="cont" />
	<input type="submit" value="�ϴ�" />
</form>
<%
	if request.TotalBytes <> 0 then
		response.Write("����ҳ�Ĵ�С��"&request.TotalBytes)
	end if
%>
</body>
</html>
