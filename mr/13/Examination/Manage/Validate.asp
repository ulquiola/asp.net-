<%
Response.CacheControl = "no-cache"
Response.Expires = 0    '����������ҳ�����л���,����֮�����̹���.
If(Session("AdminName") = "" or Session("AdminPass") = "")Then
	Response.Redirect("../index.asp")
	Response.End()
End If
%>
