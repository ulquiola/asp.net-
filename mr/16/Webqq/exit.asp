<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
a_username=Split(Application("username"),",")
username=Session("username1")
Application.Lock()
Application("username")=""
Application("username")=a_username(0)
For i=1 To UBound(a_username)
	If a_username(i)<> username Then
		Application("username")=Application("username")+","+a_username(i)
	End If
Next
Application.UnLock()
Session.Abandon()
%>
<script language="javascript">
window.close();
</script>
