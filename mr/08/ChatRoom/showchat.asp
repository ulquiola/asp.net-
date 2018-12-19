<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
response.ContentType = "text/html;charset=gb2312"
application.Lock()
A_message=split(Application("Message"),"|")
tmpChat = ""
for i=uBound(A_message) to StartV step -1
	tmpChat = tmpChat & A_message(i) & "<BR>"
Next
application.UnLock()
response.Write tmpChat
%>