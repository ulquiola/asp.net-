<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../../Conn/conn.asp"-->
<%
getproname = Request("AdminName")
getpass = Request("AdminPass")
If(getproname = "")Then
	Response.Redirect("../adm_Admin.asp")
	Response.End()
End If
Sql = "Select * from tb_Administrator"
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.Open Sql,conn,1,3
rs.AddNew
rs("Name") = getproname
rs("PWD") = getpass
rs.Update
rs.Close
Set rs = Nothing
Response.Redirect("../adm_Admin.asp")
%>
