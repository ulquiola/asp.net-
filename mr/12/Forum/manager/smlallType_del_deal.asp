<!--#include file="Conn/Conn.asp"--><!--�������ݿ������ļ�-->
<%
If Request.QueryString("ID")<>"" Then				'�жϽ��յ�IDֵ�Ƿ���ڿ�
	On Error Resume Next							'���ô�������
	SQL="Delete From tb_smalltype Where ID="&Request.QueryString("ID")	'ɾ��ָ���ļ�¼
	Conn.Execute(SQL)								'ִ��SQL���
	If Err<>0 Then						
%>
		<script language="javascript">
			alert("��Ϣɾ��ʧ�ܣ�");				//������ʾ�Ի���
			window.location.href="addsmall_manage.asp";//��ת��ָ����ASPҳ��
		</script>
	<%Else%>
		<script language="javascript">
			alert("��Ϣɾ���ɹ���");				//������ʾ�Ի���
		</script>
	<%
		response.Redirect("addsmall_manage.asp")	'��ת��ָ����ASPҳ��
	End If
End If
%>