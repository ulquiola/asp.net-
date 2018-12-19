<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../../Conn/conn.asp"-->
<%
getid = Request("id")
getname = Request("AdminName")
getpass = Request("AdminPass")
getcondition = Request("condition")
getkey = Request("key")
getpageno = Request("pageno")
Sql = "Select * from tb_Administrator Where ID = "&getid
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.Open Sql,conn,1,3
rs("Name") = getname
rs("PWD") = getpass
rs.Update
rs.Close
Set rs = Nothing
Response.Redirect("../adm_Admin.asp?condition="&getcondition&"&key="&getkey&"&pageno="&getpageno)
%>
