<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"-->
<%
if request.QueryString("id")<>"" then
id=request.QueryString("id")
end if 
set rs=server.createobject("adodb.recordset")
rs1="select * from tb_language where id="&id
rs.open rs1,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������������</title>
<style type="text/css">
<!--
.STYLE2 {font-size: 9pt}
-->
</style>
</head>

<body bgcolor="#eeeeee">
<table width="777" height="662" border="0">
  <tr>
    <td valign="top"><span class="STYLE2"><%=rs("name1")%><br>
      <%=rs("jieshi")%><br>
      <%=rs("introduce")%></span></td>
  </tr>
</table>
</body>
</html>
