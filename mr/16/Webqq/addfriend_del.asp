<!--#include file="Conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
<%
GroupID = Request.Form("GroupID") 					'���շ���IDֵ
MemberID = Request.Form("MemberID") 				'���պ��ѵ�IDֵ
 Set rs = Server.CreateObject("ADODB.Command")		'����Command����
 rs.ActiveConnection =conn 							'����Command�����������Ϣ
 sql = "delete from tb_GroupMember where id in("&MemberID&")"	'Ӧ��delete���ʵ��ɾ��
 rs.CommandText = sql								'ָ����Ҫ�����ݿ�ִ�еĲ���
 rs.Execute											'ִ��sql���
%>
<script language="javascript">
alert("ɾ���ɹ�!")										//������ʾ�Ի���
window.location = "Group_xianxi.asp?ID=<%=GroupID%>"	//��ת��ָ����ҳ��
window.parent.parent.opener.location.reload() 			//���¼���ҳ��
</script>
