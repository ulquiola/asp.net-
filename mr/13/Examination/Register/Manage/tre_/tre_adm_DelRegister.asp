<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../../Conn/conn.asp"-->
<%
getchkbox = Request("ChkBox")
getcondition = Request("condition")
getkey = Request("key")
If(getchkbox = "")Then
	Response.Redirect("../adm_Register.asp")
	Response.End()
End If
getchkbox_split=split(Replace(getchkbox," ",""),",")
For i=0 to ubound(getchkbox_split)
Sql = "Delete from tb_Student Where UserID='"&getchkbox_split(i)&"'"
conn.Execute(Sql)
Next
Response.Redirect("../adm_Register.asp?condition="&getcondition&"&key="&getkey)
%>