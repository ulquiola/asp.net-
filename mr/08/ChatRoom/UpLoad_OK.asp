<!--#include file="adovbs.inc"-->
<!--#include file="inc/func.asp"-->
<%
rbk = upfilepath("pjpeg,gif,bmp","100000")
if session("toName") <> "" then
	toName = session("toName")
else 
	toName = "������"
end if
content=session("UserName")&"<font color='#0000FF'>��</font>[<font color='#FF0000'>"&toName&"</font>]<b>������һ��Ư����ͼƬ��</b><br /><img src='selfpic/"&rbk&"' border=2 style='border-color:#f0f0f0;' > <br />("&now&")"
application.Lock()
	If application("message")="" Then
		application("message")=content
	Else
		application("message")=application("message")+"|"+content
	End If
application.UnLock()
response.Redirect("main.asp")
%>