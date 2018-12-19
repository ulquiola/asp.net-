<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"-->
<%
if request.QueryString("id")<>"" then
id=request.QueryString("id")
end if 
set pp=server.CreateObject("adodb.recordset")
qq="select * from tb_story where id="&id
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
<table width="777" height="432" border="0">
  <tr>
    <td valign="top"><span class="STYLE2"><%=replace(qq("introduce"),chr(13),"<br>")%> </span></td>
  </tr>
  <tr>
    <td height="30">
 </td>
  </tr>
</table>
</body>
</html>
