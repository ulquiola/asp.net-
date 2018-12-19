<%@ LANGUAGE = "VBScript" CodePage = "936"%>
<%
Option Explicit
Response.Buffer = True
Dim Conn,Startime
Startime = Timer()

Sub ConnectionDatabase()
	Dim ConnStr,Db
	Db = "Database/db_File.mdb"
	ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(Db)
	On Error Resume Next
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.open ConnStr
End Sub

Sub CloseDatabase()
If IsObject(Conn) then
Conn.close
Set Conn=Nothing
End If
End Sub
%>