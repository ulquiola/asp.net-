<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"-->
<%
if request.QueryString("id")<>"" then
id=request.QueryString("id")
end if 
set rs=server.createobject("adodb.recordset")
rs1="select * from tb_project where id="&id
rs.open rs1,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>明日在线搜索</title>
<style type="text/css">
<!--
.STYLE2 {font-size: 9pt}
-->
</style>
</head>
<body bgcolor="#eeeeee">
<table width="770" height="524" border="0">
  <tr>
    <td height="30" valign="top" align="center"><span class="STYLE2"><%=rs("name1")%> </span></td>
  </tr>
   <tr>
    <td height="461" valign="top"><span class="STYLE2"><%=rs("introduce")%> </span></td>
  </tr>
  <tr>
    <td height="25" valign="top">
</td>
  </tr>
</table>
</body>
</html>
