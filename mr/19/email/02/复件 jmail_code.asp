<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<% 
on error resume next
Set msg = Server.CreateObject("JMail.Message") 
msg.silent = true 
msg.Logging = true 
msg.Charset = "gb2312" 
msg.MailServerUserName =request.Form("mailfrom")  '����smtp��������֤��¼��
'msg.MailServerPassword = request.Form("mima") '����smtp��������֤���룬,������ڱ�������,����Դ��д���,��������������
msg.From =request.Form("mailfrom")  '������Email��ַ 
msg.FromName = MailServerUserName '���������� 
msg.AddRecipient(request.Form("mailto")) '�ռ���Email��ַ 
msg.Subject = request.Form("mailsubject") '�ż����� 
msg.Body = request.Form("mailbody") '����
msg.AddAttachment Server.MapPath(request.Form("file_path")),True  '����
msg.Send (request.Form("smtpaddress")) 'smtp��������ַ����ҵ�ʾֵ�ַ��
msg.close() 
set msg = nothing
If err<>0 Then
  Response.Write("<script language='JavaScript'>alert('����ʧ��,��������д��Ϣ���ٽ��в���!');history.back();</script>")
Else
  Response.Write("<script language='JavaScript'>alert('�ʼ����ͳɹ�!');window.location.href='fujian.asp';</script>")
End If
%> 
