<%
If Request.QueryString("action")="exit" Then
Session("Mname")=""
Response.Redirect("/index.asp")
End IF
If Session("Mname")="" Then
  Response.Write("<script language='javascript'>location.href='login.asp';</script>")
End If
%>
