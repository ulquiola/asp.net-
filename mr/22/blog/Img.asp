<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<%
action=Request.QueryString("action")
Select case action
case "友情链接"
  call klink()
case "相册"
  call photo()
case "相册分类"
  call photo_class()
End Select
'子过程
Sub klink()
Set rs=Server.CreateObject("ADODB.Recordset")
id=Request.QueryString("id")
sqlstr="select Kpicture from tab_klink where id="&id&""
rs.open sqlstr,conn,1,3
Response.ContentType="image/*"
Response.BinaryWrite rs("Kpicture").getChunk(8000000)
rs.close
Set rs=Nothing
End Sub
Sub photo()
Set rs=Server.CreateObject("ADODB.Recordset")
id=Request.QueryString("id")
sqlstr="select Ppic from tab_photo where id="&id&""
rs.open sqlstr,conn,1,3
Response.ContentType="image/*"
Response.BinaryWrite rs("Ppic").getChunk(8000000)
rs.close
Set rs=Nothing
End Sub
Sub photo_class()
Set rs=Server.CreateObject("ADODB.Recordset")
id=Request.QueryString("id")
sqlstr="select top 1 Ppic from tab_photo where Pclass="&id&""
rs.open sqlstr,conn,1,3
Response.ContentType="image/*"
Response.BinaryWrite rs("Ppic").getChunk(8000000)
rs.close
Set rs=Nothing
End Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>浏览图片</title>
</head>
<body>
</body>
</html>
