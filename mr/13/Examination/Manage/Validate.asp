<%
Response.CacheControl = "no-cache"
Response.Expires = 0    '把引用它的页不进行缓存,引用之后立刻过期.
If(Session("AdminName") = "" or Session("AdminPass") = "")Then
	Response.Redirect("../index.asp")
	Response.End()
End If
%>
