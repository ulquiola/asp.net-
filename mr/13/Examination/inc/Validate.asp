<%
Response.CacheControl = "no-cache"
Response.Expires = 0 
'验证访问权限
If(Session("StuID") = "" or Session("PassWord") = "")Then
	Response.Redirect("index.asp")
	Response.End()
End If
%>
