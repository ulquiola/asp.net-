<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn/conn.asp"-->
<%
if request.QueryString("id")<>"" then
id=request.QueryString("id")
end if 
set rs1=server.createobject("adodb.recordset")
rs_11="select * from tb_sousuo where id="&id
rs1.open rs_11,conn,1,3
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
<table width="777" height="538" border="0">
  <tr>
    <td height="507" valign="top"><span class="STYLE2"><%=rs1("name1")%><br>
        <%=replace(rs1("content"),chr(13),"<br>")%><br>
        <%=rs1("beizhu")%><br>
        <%=rs1("type1")%></span></td>
  </tr>
  <tr>
    <td height="25" valign="top">
    </td>
  </tr>
</table>
</body>
</html>
