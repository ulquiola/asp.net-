<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!-- #include file="Conn/Conn.asp" --><!--�������ݿ������ļ�-->
<!--�������ݿ������ļ�-->
<% 
Session.Timeout=30
function filter_Str(InString)
'***********************************
'���ܣ����������ַ����е�Σ�շ���
'���÷�����filter_Str("String")
'***********************************
	NewStr=Replace(InString,"'","''")					'��'�滻��''
	NewStr=Replace(NewStr,"<","&lt;")					'��<�滻��&lt;
	NewStr=Replace(NewStr,">","&gt;")					'��>�滻��&gt;
	NewStr=Replace(NewStr,"chr(60)","&lt;")				'��chr(60)�滻��&lt;
	NewStr=Replace(NewStr,"chr(37)","&gt;")				'��chr(37) �滻��&gt;
	NewStr=Replace(NewStr,"""","&quot;")				'��""�滻��&quot;
	NewStr=Replace(NewStr,";",";;")						'��; �滻��;;
	NewStr=Replace(NewStr,"--","-")						'��--�滻��-
	NewStr=Replace(NewStr,"/*"," ")						'��/*�滻��" "
	NewStr=Replace(NewStr,"%"," ")						'��%�滻��" "
	filter_Str=NewStr									'��NewStr�ַ�������filter_Str
end function
if Request.Form("UID")<>"" and request.Form("PWD")<>"" then								'�ж�������û����������Ƿ�Ϊ��
	UID=filter_Str(request.Form("UID"))													'��ȡ��Ԫ��UID��ֵ������UID����
	PWD=filter_Str(request.Form("PWD"))													'��ȡ��Ԫ��PWD��ֵ������PWD����	
	sql="select UserName,PWD from tb_User where UserName='"&UID&"' And Status='����Ա'"	'��ѯ����
	set rs=conn.execute(sql)							'ִ��sql���
	if rs.eof then
	%>
  		<script language="javascript">
  		alert("������Ĺ���Ա���ƴ������������룡");	//������ʾ�Ի���
	     window.location.href="index.asp";				//��ת��ָ����ASPҳ��
 		 </script> 
	<%
	Session("HS")="Manager"						'ΪSession("HS")������ֵ
	else 
		if rs("PWD")=PWD then					'�ж�������û������Ƿ���ȷ
			Session("UserName")=UID				'ΪSession("UserName")������ֵ
			Session("Status")="����Ա"			'ΪSession("Status")������ֵ
	lastlogintime=now()							'��ȡ��ǰϵͳ�����ں�ʱ��
	UP="Update tb_user set logintimes=logintimes+1,lastlogintime='"&lastlogintime&"' where UserName='"&UID&"' And Status='����Ա'"
	'Ӧ��Update������ָ���ļ�¼
	conn.execute(UP)								'ִ��UP���
	%>
<script language=javascript>
 alert("����Ա��¼�ɹ���");							//������ʾ�Ի���
//parent.leftwindow.location.reload();				//ˢ��ָ����ҳ��
parent.location.href="manage_main.asp";				//��ת��ָ����ASPҳ��
</script>

		<%else%>
   			<script language="javascript">
			 alert("������Ĺ���Ա����������������룡");			//������ʾ�Ի���
  			 window.location.href="index.asp";						//��ת��ָ����ASPҳ��
  			</script>  		
		<%
		Session("HS")="Manager"										'ΪSession("HS")������ֵ
		end if
	end if
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">