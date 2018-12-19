<%
response.BinaryWrite(Application("Send_img"))
Application.Lock()
Application("Send_img")=""
Application.UnLock()
%>