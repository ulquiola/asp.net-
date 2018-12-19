<%@ Language=VBScript %>
<%
userID = ",S"&Request.querystring("UserID")&"E"
application("OnLineUserID") = replace(application("OnLineUserID"),UserID,"")
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script>
window.close()
</script>
</HEAD>
<BODY>

<P>&nbsp;</P>
</BODY>
</HTML>
