<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
If Not Isempty(Request("send")) Then 
on error resume next
Set msg = Server.CreateObject("JMail.Message") 
msg.silent = true 
msg.Logging = true 
'msg.Charset = "gb2312" 
msg.MailServerUserName =request.Form("mailfrom")  '����smtp��������֤��¼��,���κ�һ���ڸ�smpt�������������Email��ַ�� 
'msg.MailServerPassword = request.Form("mima")	  '����smtp��������֤����,��Email�˺Ŷ�Ӧ������,������ڱ�������,����Դ��д���,��������������
msg.From =request.Form("mailfrom")  '������Email��ַ 
msg.FromName = MailServerUserName  '���������� 
msg.AddRecipient(request.Form("emails")) '�ռ���Email��ַ 
msg.Subject = request.Form("mailsubject") '�ż����� 
msg.Body = request.Form("mailbody") '���� 
msg.Send (request.Form("smtpaddress")) 'smtp��������ַ����ҵ�ʾֵ�ַ��
msg.close() 
set msg = nothing 
If err<>0 Then
  Response.Write("<script language='JavaScript'>alert('�ʼ�����ʧ��,���ʵ���������Ƿ�׼ȷ!');history.back();</script>")
Else
  Response.Write("<script language='JavaScript'>alert('ʹ��Jmail��������ʼ��ɹ�!');window.location.href='jmail.asp';</script>")
End If
End If
%> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ʹ��Jmail��������ʼ�</title>
<style>
.style1 {color: #FF0000}
body {
	margin-left: 0px;
	margin-top: 0px;
}
body,td,th {
	font-size: 12px;
}
</style>
<script type="text/javascript">
function Mycheck(form){
  for(i=0;i<form.length;i++){
    if(form.elements[i].value==""){
	  alert("������������Ϣ����!");return false;}		
  }
}
</script>
</head>
<body style="font-size:12px">
<table width="755" height="365" border="0" cellspacing="0" cellpadding="0">
<form name="form_jamil" action="jmail.asp" method="post">
  <tr>
  	<td height="365" align="left" valign="top" background="images/bg1.jpg" id="peop"><table width="755" height="365"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="17%" align="right" valign="middle">�ռ��ˣ�</td>
        <td width="5%" align="left" valign="middle">&nbsp;</td>
        <td width="78%" align="left" valign="middle">
          <input name="emails" type="text" id="emails" size="40" value="<%=Request.QueryString("mailto")%>">      
          <span class="style1">*</span></td>
      </tr>
      <tr>
        <td align="right" valign="middle">�����ˣ�</td>
        <td align="left" valign="middle">&nbsp;</td>
        <td align="left" valign="middle"><span class="style4">
      <input name="mailfrom" type="text" id="mailfrom" size="40">
      <span class="style1">*</span> </span></td>
      </tr>
      <tr>
        <td align="right" valign="middle">���룺</td>
        <td align="left" valign="middle">&nbsp;</td>
        <td align="left" valign="middle"><span class="style4">
      <input name="mima" type="password" id="mima" size="44">
      <span class="style1">*</span> </span></td>
      </tr>
      <tr>
        <td align="right" valign="middle">SMTP��������ַ��</td>
        <td align="left" valign="middle">&nbsp;</td>
        <td align="left" valign="middle"><span class="style4">
      <input name="smtpaddress" type="text" id="smtpaddress" size="40">
    (��:smtp.163.com��)<span class="style1">*</span></span></td>
      </tr>
      <tr>
        <td align="right" valign="middle">���⣺</td>
        <td align="left" valign="middle">&nbsp;</td>
        <td align="left" valign="middle"><span class="style4">
      <input name="mailsubject" type="text" id="mailsubject" size="40">
      <span class="style1">*</span> </span></td>
      </tr>
      <tr>
        <td align="right" valign="middle">���ݣ�</td>
        <td align="left" valign="middle"><span>
        </span></td>
        <td align="left" valign="middle"><span class="style4">
      <textarea name="mailbody" cols="43" rows="6" wrap="PHYSICAL" id="mailbody"></textarea>
      <span class="style1">*</span>    </span></td>
      </tr>
      <tr>
        <td align="right" valign="middle">&nbsp;</td>
        <td align="center" valign="middle">&nbsp;</td>
        <td align="left" valign="middle"><input name="send" type="submit" id="send" value="����" onClick="return Mycheck(this.form)">
��
  <input type="reset" name="Submit2" value="��д"></td>
      </tr>
      <tr>
        <td colspan="3" align="center" valign="middle"><span>*�������ݱ�����д</span></td>
        </tr>
    </table></td>
  </tr>
  </form>
</table>
<table width="755" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="42" background="images/bottom.jpg"></td>
  </tr>
</table>
</body>
</html>
