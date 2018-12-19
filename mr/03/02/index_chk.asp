<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
</head>
<body>
&nbsp;&nbsp;&nbsp;&nbsp;会员信息<br>
会员昵称：<%=request.Form("name1")%><br> 	<%'应用Form数据集合获取表单数据%>
会员名称：<%=request.Form("name2")%><br>	<%'应用Form数据集合获取表单数据%>
所属地区：<%=request.Form("dq")%><br> 		<%'应用Form数据集合获取表单数据%>
通讯地址：<%=request.Form("address")%> 		<%'应用Form数据集合获取表单数据%>

</body>
</html>
