<!--#include file="../config.asp"-->
<% 
'连接数据库
dim driver,conn
	driver = "provider=sqloledb;data source=("&URL&");initial catalog="&DataName&";user id="&User&";password="&Password&";"
	On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open driver
	If Err Then
		Err.Clear
		conn.Close
		Set conn = Nothing
		Response.Write "系统数据库连接错误!"
		Response.End
	else
	End If
	response.ContentType = "text/html; charset=GBK"
%>



