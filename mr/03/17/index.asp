<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>MapPath方法</title>
<style>
body{
	margin:12px;
	font-size:12px;
}
</style>
</head>
<body>
<%
Response.Write("服务器宿主目录的物理路径为："&Server.MapPath("\")&"<BR>")
Response.Write("当前物理路径为："&Server.MapPath("./")&"<BR>")
Response.Write("父目录物理路径为："&Server.MapPath("../")&"<BR>")
Response.Write("当前文件物理路径为："&Server.MapPath(Request.ServerVariables("PATH_INFO"))) 
'使用Request对象的服务器变量PATH_INFO来映射当前文件的物理路径
%>

</body>
</html>
