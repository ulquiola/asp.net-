<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
If Not Isempty(Request("sure")) Then 
Emails=Split(Request("email"),",") 
for i=0 to ubound(Emails)
Set msg = Server.CreateObject("JMail.Message") '���ܷ���ѭ������
msg.silent = true 
msg.Logging = true 
msg.Charset = "gb2312" 
msg.MailServerUserName =request.Form("mailfrom")  '����smtp��������֤��¼��,���κ�һ���ڸ�smpt�������������Email��ַ�� 
'msg.MailServerPassword = request.Form("mima") '����smtp��������֤����,��Email�˺Ŷ�Ӧ������,������ڱ�������,����Դ��д���,��������������
msg.From =request.Form("mailfrom")  '������Email��ַ 
msg.FromName = "pyj" '���������� 
msg.AddRecipient(Emails(i)) '�ռ���Email��ַ 
msg.Subject = request.Form("mailsubject") '�ż����� 
msg.Body = request.Form("mailbody") '���� 
msg.Send (request.Form("smtpaddress")) 'smtp��������ַ����ҵ�ʾֵ�ַ��
msg.close()
set msg = nothing 
next
Response.Write("<script language='JavaScript'>alert('�ʼ�Ⱥ���ɹ�!');window.location.href='qunfa.asp';</script>")
End IF
%>

<script language="javascript">
function CheckAll(elementsA){
	var len=elementsA;
		for (i=0;i<len.length;i++){
			elementsA[i].checked=false;
			}
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ʵ���ʼ�Ⱥ��</title>
<!--#include file="conn.asp"-->
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
body,td,th {
	font-size: 12px;
}
-->
</style></head>
<body>
<%
	set rs = server.CreateObject("adodb.recordset")
	sql = "select * from jmail"
	rs.open sql,conn,1,1
%>
<table width="755" height="400" border="0" align="center" cellpadding="0" cellspacing="0" background="images/bg2.jpg">
  <form action="" method="post" name="myform">
  <tr>
    <td height="30" colspan="4"  align="center" valign="middle">ͨѶ¼</td>
  </tr>
<tr>
  <td align="left" valign="top" id="peop">
  <table width="90%" height="153" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#99CC00">
  <tr>
    <td width="104" height="30"><div align="center">ѡ��</div></td>
    <td width="60"><div align="center">���</div></td>
    <td width="100"><div align="center">�ռ��˵�����</div></td>
    <td width="309" align="center" valign="middle"><div>�ռ��˵�����</div></td>
  </tr>
  <%
  	do while not rs.eof
  %>
  <tr>
  	<td height="25"><div align="center" ><input type="checkbox" name="email" value=<%=rs("email")%>></div></td>
    <td height="25"><div align="center"><%=rs("id")%></div></td>
    <td height="25"><div align="center"><%=rs("name")%></div></td>
    <td height="25" align="center" valign="middle"><a href="jmail.asp?mailto=<%=rs("email")%>"><%=rs("email")%></a></td>
  </tr>
<%
	rs.movenext
	loop
	rs.close
	conn.close
%>
  <tr align="left" valign="middle">
  	<td height="30" align="center" class="style1"><input type="button" name="Submit4" value="ȡ��ѡ��" onClick="CheckAll(myform.email)"></td>
  	<td height="30" colspan="3" class="style1" style="line-height:18px">ע�⣺ѡ�и�ѡ�򣬿�������ʼ����ռ��ˣ�������ȡ��ѡ�񡱰�ť����ȡ������ѡ������½���ѡ��</td>
  	</tr>
 </table>
 </td>
</tr>
  <tr>
  	<td height="10" colspan="4" align="center" valign="top" id="peop">
		<table width="755" border="0" align="center">
  			<tr>
    			<td width="127" height="22" align="right" valign="middle"><span>�����ˣ�</span></td>
    			<td width="33" align="right" valign="middle">&nbsp;</td>
    			<td width="581" height="22" align="left" valign="middle"><span><input name="mailfrom" type="text" id="mailfrom" size="40"></span></td>
 			</tr>
  			<tr>
    			<td height="22" align="right" valign="middle"><span>���룺</span></td>
    			<td height="22" align="right" valign="middle">&nbsp;</td>
    			<td height="22" align="left" valign="middle"><span><input name="mima" type="password" id="mima" size="44"></span></td>
  			</tr>
  			<tr>
    			<td height="22" align="right" valign="middle"><span>SMTP��������ַ��    </span></td>
    			<td height="22" align="right" valign="middle">&nbsp;</td>
    			<td height="22" align="left" valign="middle"><span ><input name="smtpaddress" type="text" id="smtpaddress" size="40"></span></td>
  			</tr>
  			<tr>
   				 <td height="22" align="right" valign="middle"><span>���⣺</span></td>
    			 <td height="22" align="right" valign="middle">&nbsp;</td>
   			    <td height="22" align="left" valign="middle"><span><input name="mailsubject" type="text" id="mailsubject" size="40"></span></td>
  			</tr>
  			<tr>
    			<td align="right" valign="middle"><span>�����ݣ�</span></td>
   				<td align="right" valign="middle">&nbsp;</td>
   				<td align="left" valign="middle"><span ><textarea name="mailbody" cols="40" rows="5" wrap="PHYSICAL" id="mailbody"></textarea></span></td>
  			</tr>
  			<tr>
   				<td height="30" align="right" valign="middle">&nbsp;</td>
  			    <td height="30" align="center" valign="middle">&nbsp;</td>
  			    <td height="30" align="left" valign="middle"><input name="sure" type="submit" id="sure" value="����">
  			      &nbsp;&nbsp;
		        <input type="reset" name="Submit2" value="��д"></td>
  			</tr>
		</table>
	</td>
  </tr>
  </form>
</table>
<table width="755" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="42" background="images/bottom.jpg"></td>
	</tr>
</table>
