<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"-->
<%
if request.querystring("ID")<>"" then
ID=request.querystring("ID")
end if 
set ab=server.CreateObject("adodb.recordset")
cc="select * from tb_project  where ID="&ID 
ab.open cc,conn,1,3
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
<table width="777" height="519" border="0">
  <tr>
    <td height="490" valign="top"><p align="left" class="STYLE1"><%=ab("introduce")%> </p>   </td>
  </tr>
  <tr>
    <td height="23"><p align="left">　　　　　　　　　　　　　　　　　　　　　
        <input name="myclose" type="button" class="Style_button_del" id="myclose"
				 value="关闭窗口" onClick="javascrip:opener.parent.location.reload();self.close()">
                                        </p></td>
  </tr>
</table>
</body>
</html>
