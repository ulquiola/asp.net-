<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="Conn/Conn.asp"-->
<%
If Request.QueryString("TopicID")<>"" Then
	On Error Resume Next
	SQL="Delete From tb_Topic Where ID="&Request.QueryString("TopicID")
	Conn.Execute(SQL)
	If Err<>0 Then
%>
		<script language="javascript">alert("������Ϣɾ��ʧ�ܣ�");history.back(-1);</script>
	<%Else%>
		<script language="javascript">alert("������Ϣɾ���ɹ���");window.close();</script>
	<%End If
End If
%>
