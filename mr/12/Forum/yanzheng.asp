<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
typeid=request.querystring("TypeID")					'��ȡ��Ԫ��ID��ֵ������ID����
If Session("UserName")="" or Session("UserName")="Friend" Then%>
<script language="javascript">
alert("��ע���Ϊ�û����ſ��Է���������Ϣ��");
window.location.href="register.asp";
</script>
<%else%>
<script language="javascript">
window.location.href="Add_Topic.asp?typeid=<%=typeid%>";
</script>
<%end if%>
<%If Session("UserName")="" or Session("UserName")="Friend"  Then%>
<a href="#" onClick="Myopen(User,form_U)">�û���¼</a>
<%Else%>
<a href="Modify.asp"> �޸�����</a>
<%End If%> 
�� <a href="index.asp">�鿴�������</a> �� <a href="#" onClick="window.location.reload();">ˢ��ҳ��</a> �� 
<%If Session("UserName")="" or Session("UserName")="Friend"  Then%>
<a href="#" onClick="Myopen(Manager,form_M)">����Ա��¼</a> �� <a href="friend_Add_Topic.asp">��������</a> 
<%Else%>
<a href="Logout.asp"> ע���û�</a>
<%End If%>