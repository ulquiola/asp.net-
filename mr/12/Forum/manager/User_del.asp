<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="Conn/Conn.asp"--><!--�������ݿ������ļ�-->
<%
If Request.QueryString("UID")<>"" Then					'�ж�UID�Ƿ���ֵ
	Set rs=Server.CreateObject("ADODB.RecordSet")		'������¼��
	SQL="Select * From tb_Topic Where Author="&Request.QueryString("UID")	'��ѯ����
	rs.Open SQL,conn,1,3													'�򿪼�¼��
	If Not(rs.Eof and rs.Bof) Then											'�ж��Ƿ��м�¼��Ϣ
%>
		<script language="javascript">
		if(!confirm("���û��Ѿ����������Ҫɾ�����û���")){//������ʾ�Ի���
			window.location.href="User_Manage.asp";//��ת��ָ����ASPҳ��
		}
		</script>
	<%
	End If
		On Error Resume Next											'���ô�������
		SQL="Delete From tb_User Where ID="&Request.QueryString("UID")'ɾ��ָ�����û���Ϣ
		Conn.Execute(SQL)												'ִ��SQL���
		If Err<>0 Then
	%>
			<script language="javascript">
				alert("�û�ɾ��ʧ�ܣ�");					//������ʾ�Ի���
				window.location.href="User_Manage.asp";		//��ת��ָ����ASPҳ��
			</script>
		<%Else%>
			<script language="javascript">
				alert("�û�ɾ���ɹ���");				//������ʾ�Ի���
				window.location.href="User_Manage.asp";	//��ת��ָ����ASPҳ��
			</script>
		<%End If
End If
%>