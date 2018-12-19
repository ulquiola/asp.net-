<%@LANGUAGE="VBSCRIPT"%>
<%
MR_logoutRedirectPage = "logout.asp"
Session.Abandon
If (MR_logoutRedirectPage <> "") Then 
Response.Redirect(MR_logoutRedirectPage)
end if
%>
