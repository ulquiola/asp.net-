<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="Conn/Conn.asp"--><!--�������ݿ������ļ�-->
<%
ID=Request.QueryString("ID")							'��ȡID��ֵ
If ID<>"" Then											'�ж�ID��ֵ�Ƿ��
	Set rs=Server.CreateObject("ADODB.RecordSet")		'������¼��
	SQL="Select * From tb_smalltype Where ID="&ID		'��ѯ����
	rs.Open SQL,conn,1,3								'�򿪼�¼��
	If Not(rs.Eof and rs.Bof) Then'�ж��Ƿ��м�¼��Ϣ%>					
		<script language="javascript">
		if(!confirm("�Ѿ����ڸð������������Ϣ��\n���ɾ���ð�飬��ʱ������Ҳ��һͬɾ���������Ҫɾ���ð����")){
			window.location.href="addsmall_manage.asp";//������ʾ�Ի���
		}else{
		window.location.href="smlallType_del_deal.asp?ID=<%=ID%>";//��ת��ָ����ASPҳ��
		}
		</script>
	<%
	Else
		response.Redirect("smlallType_del_deal.asp?ID="&ID)			'��ת��ָ����ASPҳ��
	End IF
	
End If
%>