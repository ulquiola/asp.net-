<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
<%
UserID =split(Request.QueryString("ID"),",")		'Ӧ��split�������ػ���0��һά���飬���а���ָ����Ŀ�����ַ���
GroupID = Request.QueryString("GroupID")			'��ȡ�������Ƶ�ID��
for i=0 to ubound(UserID)							'Ӧ��ubound������ȡ�������±�
	Set rs = Server.CreateObject("ADODB.RecordSet")	'������¼��
	rs.ActiveConnection = conn 						'��¼��
	rs.Source = "SELECT  * FROM tb_groupMember WHERE UserID="&UserID(i)&" and GroupID="&GroupID&""	'������Ϣ��ѯ
	rs.CursorType = 0
	rs.CursorLocation = 2
	rs.LockType =2
	rs.Open()
	if rs.EOF then
		rs.AddNew 									'Ӧ��AddNew����ʵ���������
		rs.fields("UserID").value = UserID(i) 
		rs.fields("GroupID").value = GroupID
		rs.Update 
	end if
next
%>
<script language="javascript">
window.location="Group_xianxi.asp?ID=<%=GroupID%>"		//��ת�������б�ҳ��
window.parent.parent.opener.location.reload() 			//ˢ�¸�����
window.close();											//�û������ɺ󣬹ر�ָ���Ĵ���
</script>
