<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="Conn/Conn.asp"-->
<%
ReplyID=Request.QueryString("ReplyID")
If ReplyID<>"" Then
	On Error Resume Next
	SQL="Delete From tb_Reply Where ID="&ReplyID
	Conn.Execute(SQL)
	If Err<>0 Then
%>
		<script language="javascript">
			alert("�ظ���Ϣɾ��ʧ�ܣ�");
			parent.location.href="show.asp";
		</script>
	<%Else%>
		<script language="javascript">
			alert("�ظ���Ϣɾ���ɹ���");
			parent.location.href="show.asp";
		</script>
	<%
End If
End If
%>
