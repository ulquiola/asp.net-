<!--#include file="../config.asp"-->
<% 
'�������ݿ�
dim driver,conn
	driver = "provider=sqloledb;data source=("&URL&");initial catalog="&DataName&";user id="&User&";password="&Password&";"
	On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open driver
	If Err Then
		Err.Clear
		conn.Close
		Set conn = Nothing
		Response.Write "ϵͳ���ݿ����Ӵ���!"
		Response.End
	else
	End If
	response.ContentType = "text/html; charset=GBK"
%>



