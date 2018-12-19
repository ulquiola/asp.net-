<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
  Dim Conn,Str									'定义变量
  Set Conn=Server.CreateObject("ADODB.Connection") 		'创建Connection对象实例
  Str="Driver={Microsoft Access Driver (*.mdb)};DBQ="&Server.mappath("DataBase/db_8.mdb")&"" 
'定义连接数据库字符串
  Conn.Open(Str) 									'建立连接
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>读取结果集下的字段</title>
<style>
body{
	margin:12px;
}
body,th,td{
	font-size:12px;
}
</style>
</head>
<body>
<%
set rs=server.CreateObject("adodb.recordset")				'创建recordset对象实例
sql="select * from tb_user"							'建立连接
rs.open sql,conn,1,3									'打开记录集
%>
<table width="437" height="56" border="1" align="center" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#CCCCCC">
  <tr>
    <td bgcolor="#FFCC33"><div align="center">用户名称</div></td>
    <td bgcolor="#FFCC33"><div align="center">用户密码</div></td>
    <td bgcolor="#FFCC33"><div align="center">通讯地址</div></td>
    <td bgcolor="#FFCC33"><div align="center">邮政编码</div></td>
  </tr>
  <tr>
    <td><div align="center"><%=Rs.Fields("uname").value%></div></td><!--获取字段信息-->
    <td><div align="center"><%=Rs.Fields("pwd").value%></div></td><!--获取字段信息-->
    <td><div align="center"><%=Rs.Fields("address").value%></div></td><!--获取字段信息-->
    <td><div align="center"><%=Rs.Fields("postcode").value%></div></td><!--获取字段信息-->
  </tr>
</table>

</body>
</html>
