<!--#include file="conn/conn.asp"-->
<%
if session("UserName")="" then			'�ж��û����Ƿ�Ϊ��
%>
<% 
		session.Timeout=120				'����Session�����ʹ��ʱ��
	if request.Form("UserName")<>"" and request.Form("PWD")<>"" then	'�ж��û����������Ƿ�Ϊ��
		session("UserName")=request.Form("UserName")					'����ȡ���û����Ƹ�ֵ��Session����
		session("PWD")=request.Form("PWD")								'����ȡ���û����븳ֵ��Session����
		sql="select UserName,PWD from tb_user where UserName='"&session("UserName")&"'"	'��ѯ����
		set rs=conn.execute(sql)														'ִ��SQL���
		if rs.eof then 
	%>
			<script language="javascript">
			alert("��������û����������������룡");								//������ʾ�Ի���
			 </script> 
			 <%
			 session.Abandon()  
			 'ɾ�����д���Session�����еĶ���
		else 
			if rs("PWD")=session("PWD") then									'�ж��û������Ƿ����
				session("PWD")=request.Form("pwd") '���û����븳ֵ��Session����%>							
				<script language="javascript">
				window.location.href="index.asp"					//��ת��ָ����ҳ��
				</script>
			 <%else%>
				<script language="javascript">
				 alert("��������û�����������������룡");		//������ʾ�Ի���
				</script>  		
				<%
				session.Abandon()
			end if
		end if
end if
%>
	<script language="javascript">
	function Mycheck()							//�����Զ��庯��
	{
	if(form1.username.value=="")				//�ж��û����Ƶ��ı����ֵ�Ƿ�Ϊ��
	{alert("�������û������ٽ��е�¼!");form1.username.focus();return;}
	if(form1.pwd.value=="")						//�ж��û�������ı����ֵ�Ƿ�Ϊ��
	{alert("�û����벻��Ϊ��!");form1.pwd.focus();return;}
	form1.submit();								//���б��ύ
	}
	</script>
	<style type="text/css">
<!--
.STYLE4 {
font-size: 9pt; color: #000000;
border:1px solid #B3B3B3;
}
.STYLE5 {
	font-size: 9pt;
	font-weight: bold;
	color: #606060;
}
-->
    </style>
	<style type="text/css">
<!--
body {
	margin-top: 0px;
	margin-bottom: 0px;
}
.STYLE6 {
	font-size: 10pt;
	color: #FF0000;
	font-weight: bold;
}

-->
    </style>
	<table width="217" height="109" border="0" align="center" cellpadding="0" cellspacing="0">
	  <tr>
		<td width="169" height="109">
		<form action="index.asp" name="form1" method="post">
				<table width="204" height="67" border="0" align="left" cellpadding="0" cellspacing="0">
				  <tr>
					<td width="68" height="28" valign="bottom"><div align="center">&nbsp;&nbsp;&nbsp;�û���</div></td>
					<td width="105" valign="bottom"><input name="username" type="text" class="STYLE4" id="username" size="15"></td>
				  </tr>
				  <tr>
					<td height="18"><div align="center">�� ��</div></td>
					<td><input name="pwd" type="password" class="STYLE4" id="pwd" size="16" /></td>
				  </tr>
				  <tr>
					<td height="21" colspan="2" valign="bottom"><div align="center"><br>
					  &nbsp;
					  <input type="button" name="Submit2" value="��¼" onClick="Mycheck();">
					  &nbsp;
					  <input type="reset" name="Submit3" value="����">
					</div></td>
				  </tr>
	      </table>
<br><br><br><br><br><br>	            <p align="right" class="STYLE5" onClick="JScript:window.location.href='reg.asp';">�û�ע��</p>
		</form></td>
    </tr>
	</table>
	<%else%><table width="171" height="42" border="0" align="center" cellpadding="0" cellspacing="0">
			<tr>
			  <td width="66" valign="bottom"><div align="right" class="STYLE6"><%=session("UserName")%></div></td>
			  <td width="128" valign="bottom"><span class="STYLE4">&nbsp;��ӭ�����ٱ���վ</span></td>
			</tr>
			<tr>
			  <td height="21" colspan="2"><div align="center">
				  <input name="button1" type="button" value="�˳�" onClick="if(confirm('�Ƿ�����˳�?')){window.location.href='exit.asp'}">
			  </div></td>
			</tr>
		  </table>
	<%end if%>
