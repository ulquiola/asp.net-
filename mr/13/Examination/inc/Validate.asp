<%
Response.CacheControl = "no-cache"
Response.Expires = 0 
'��֤����Ȩ��
If(Session("StuID") = "" or Session("PassWord") = "")Then
	Response.Redirect("index.asp")
	Response.End()
End If
%>
