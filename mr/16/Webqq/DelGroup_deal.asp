<!--#include file="Conn/conn.asp" -->
<%'���ݿ������ļ�%>
<!--#include file="js/Function.asp" -->
<%
GroupID = Request.Form("GroupID")							'����GroupID��ֵ
 Set rs = Server.CreateObject("ADODB.Command")				'����Command����
 rs.ActiveConnection =conn 									'����Command�����������Ϣ
 sql = "delete from tb_GroupMain where id in("&GroupID&")"	'Ӧ��delete���ʵ�ּ�¼��ɾ��
 rs.CommandText = sql										'ָ����Ҫ�����ݿ�ִ�еĲ���
 rs.Execute													'ִ��sql���
%>
<script>
alert("ɾ���ɹ�!")											//������ʾ�Ի���
window.location = "GList.asp"								//��ת��ָ����ҳ��
window.parent.parent.opener.location.reload() 				//���¼���ҳ��
</script>