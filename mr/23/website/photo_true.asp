<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Conn/conn.asp" -->
<%
Dim ps__MRColParam
ps__MRColParam = "1"
if (Request.QueryString("idd") <> "") then ps__MRColParam = Request.QueryString("idd")
%>
<%
set ps = Server.CreateObject("ADODB.Recordset")
ps.ActiveConnection = MR_website_STRING
ps.Source = "SELECT * FROM photos WHERE idd = " + Replace(ps__MRColParam, "'", "''") + ""
ps.CursorType = 0
ps.CursorLocation = 2
ps.LockType = 3
ps.Open()
ps_numRows = 0
%>
<html>
<head>
<title>图片实际大小</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link 
href="css/css.css" rel=stylesheet type="text/css">
<style type="text/css">
<!--
.pt9 {  font-family: "宋体"; font-size: 9pt}
-->
</style>
</head>
<body bgcolor="#EFEFEF" text="#000000" leftmargin="0" topmargin="0">
<table border="0" cellspacing="0" cellpadding="0" class="pt9" width="100%">
  <tr>
    <td height="474"><div align="center"><img src="../website/photos/<%=(ps.Fields.Item("tp").Value)%>"></div></td>
  </tr>
  <tr>
    <td height="30"><div align="center"><b>[<a href="" onClick="window.close();return false;"">关闭窗口</a>]</b></div></td>
  </tr>
</table>
</body>
</html>
<%
ps.Close()
%>
