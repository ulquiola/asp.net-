<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"-->
<%
if request.QueryString("id")<>"" then
id=request.QueryString("id")
end if 
set rs1=server.createobject("adodb.recordset")
rs_11="select * from tb_story where id="&id
rs1.open rs_11,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>明日在线搜索</title>
<style type="text/css">
<!--
.STYLE1 {font-size: 9pt}
-->
</style>
</head>
<body bgcolor="#eeeeee">
<table width="777" height="538" border="0">
  <tr>
    <td height="507" valign="top"><div align="center" class="STYLE1"><%=rs1("title")%><br>
        <%=replace(rs1("content"),chr(13),"<br>")%></div></td>
  </tr>
  <tr>
    <td height="25" valign="top">
    </td>
  </tr>
</table>
</body>
</html>
