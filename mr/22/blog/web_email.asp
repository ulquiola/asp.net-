<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
If Not Isempty(Request("send")) Then 
Set msg = Server.CreateObject("JMail.Message") 
msg.silent = true 
msg.Logging = true 
msg.Charset = "gb2312" 
msg.MailServerUserName =request("mailfrom")  '����smtp��������֤��½�� ���ʾ����κ�һ���û���Email��ַ�� 
msg.MailServerPassword = request("mima") '����smtp��������֤���� ���û�Email�ʺŶ�Ӧ�����룩 
msg.From =request("mailfrom")  '������Email��ַ 
msg.FromName = "pyj" '���������� 
msg.AddRecipient(request("emails")) '�ռ���Email��ַ 
msg.Subject = request("mailsubject") '�ż����� 
msg.Body = request("mailbody") '���� 
msg.Send (request("smtpaddress")) '����smtp��������ַ����ҵ�ʾֵ�ַ�� 
set msg = nothing 
response.Write "<br>"
response.Write "���ͳɹ�����"
End If
%> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����----�����ʼ�</title>
</head>
<body>
<table width="500" height="281" border="0" align="center">
  <tr>
    <td height="29" colspan="2"  align="center" valign="middle"><p align="center">�����ż�����</p></td>
  </tr>
  <tr>
    <td width="127" height="22" align="right" valign="middle"><span class="style4">�ռ���email�� </span></td>
    <td width="363" align="left" valign="middle"><span class="style4">
      <input name="emails" type="text" id="emails">
    </span></td>
  </tr>
  <tr>
    <td height="22" align="right" valign="middle"><span class="style4">������email�� </span></td>
    <td align="left" valign="middle"><span class="style4">
      <input name="mailfrom" type="text" id="mailfrom">
    </span></td>
  </tr>
  <tr>
    <td height="18" align="right" valign="middle"><span class="style4">���룺 </span></td>
    <td height="18" align="left" valign="middle"><span class="style4">
      <input name="mima" type="password" id="mima">
    </span></td>
  </tr>
  <tr>
    <td height="21" align="right" valign="middle"><span class="style4">SMTP��������ַ�� </span></td>
    <td height="21" align="left" valign="middle"><span class="style4">
      <input name="smtpaddress" type="text" id="smtpaddress">
    </span></td>
  </tr>
  <tr>
    <td height="18" align="right" valign="middle"><span class="style4">���⣺�� </span></td>
    <td height="18" align="left" valign="middle"><span class="style4">
      <input name="mailsubject" type="text" id="mailsubject">
    </span></td>
  </tr>
  <tr>
    <td height="83" align="right" valign="middle"><span class="style4">�����ݣ��� </span></td>
    <td height="83" align="left" valign="middle"><span class="style4">
      <textarea name="mailbody" cols="45" rows="6" wrap="PHYSICAL" id="mailbody"></textarea>
    </span></td>
  </tr>
  <tr>
    <td height="37" colspan="2" align="center" valign="middle">
      <input name="send" type="submit" id="send" value="����">
&nbsp;
    <input type="reset" name="Submit2" value="��д">    </td>
  </tr>
</table>
</body>
</html>
