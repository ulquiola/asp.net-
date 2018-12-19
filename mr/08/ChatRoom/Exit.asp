<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
A_UserName=Split(Application("UserName"),",")
A_head=Split(Application("head"),",")
A_IP=Split(Application("UserIP"),",")
UserName=Session("UserName")
Application.Lock()
Application("UserName")=" "
Application("head")=" "
Application("UserName")=A_UserName(0)
Application("head")=A_head(0)
For i=1 To UBound(A_UserName)
	If A_UserName(i)<> UserName Then
		Application("UserName")=Application("UserName")+","+A_UserName(i)
		Application("head")=Application("head")+","+A_head(i)
		Application("UserIP")=Application("UserIP")+","+A_IP(i)
	End If
Next
Application.UnLock()
Session.Abandon()
%>
<script language="javascript">
window.close();
</script>