<!--#include file="conn/conn.asp" -->
<%'���ݿ������ļ�%>
<!--#include file="js/Function.asp" -->
<%'����������֤��ҳ��%>
<%
Set rs = Server.CreateObject("ADODB.RecordSet")				'������¼��
rs.ActiveConnection = conn 									'
rs.Open "tb_talk",,1,3										'������Դ
rs.AddNew 
rs.fields("fuserid").value = request.form("fuserid")				'������Ϣ�û���ID
rs.fields("fusernickname").value = request.form("fusernickname")	'������Ϣ�û����ǳ�
rs.fields("touserid").value = request.form("touserid") 				'������Ϣ���û�ID
rs.fields("message").value = request.form("Message")				'��ȡ������Ϣ
rs.Update 															'�������ݸ���
%>
<script language="javascript">
alert('��Ϣ���ͳɹ�');												//������ʾ��Ϣ
window.close();														//�رյ���ҳ��
</script>