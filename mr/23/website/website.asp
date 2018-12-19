<!--#include file="Conn/conn.asp" -->
<%
Dim Recordset1__MRColParam
Recordset1__MRColParam = "1"
if (Request.QueryString("id") <> "") then Recordset1__MRColParam = Request.QueryString("id")
%>
<%
set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MR_website_STRING
Recordset1.Source = "SELECT * FROM website WHERE id = '" + Replace(Recordset1__MRColParam, "'", "''") + "'"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 3
Recordset1.Open()
Recordset1_numRows = 0
%>
<html>
<head>
<title>自动建站</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="refresh" content="0;URL=../templet<%=(Recordset1.Fields.Item("xuanze").Value)%>/page1.asp?id=<%=(Recordset1.Fields.Item("id").Value)%>">
</head>
<body bgcolor="#FFFFFF" text="#000000">
</body>
</html>
<%
Recordset1.Close()
%>
