<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="Conn/Conn.asp"--><!--�������ݿ������ļ�-->
<%
QID=Request.QueryString("ID")							'ΪQID������ֵ
If QID<>"" Then											'�ж�IDֵ�Ƿ���ڿ�
	Set rs=Server.CreateObject("ADODB.RecordSet")		'������¼��
	SQL="Select * From tb_Topic Where TypeID="&QID		'��ѯ����
	rs.Open SQL,conn,1,3								'�򿪼�¼��
	If Not(rs.Eof and rs.Bof) Then'�Ƿ��м�¼��Ϣ%>					
		<script language="javascript">
		if(!confirm("�Ѿ����ڸ���������������Ϣ��\n���ɾ�����������������Ҳ��һͬɾ���������Ҫɾ���������")){
			window.location.href="Type_Manage.asp";
			//��ת��ָ����ASPҳ��
		}else{
		window.location.href="Type_del_deal.asp?ID=<%=QID%>";//��ת��ָ����ASPҳ��
		}
		</script>
	<%
	Else
		response.Redirect("Type_del_deal.asp?ID="&QID)		'��ת��ָ����ASPҳ��
	End IF
	
End If
%>