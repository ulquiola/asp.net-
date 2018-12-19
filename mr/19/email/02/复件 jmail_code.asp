<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<% 
on error resume next
Set msg = Server.CreateObject("JMail.Message") 
msg.silent = true 
msg.Logging = true 
msg.Charset = "gb2312" 
msg.MailServerUserName =request.Form("mailfrom")  '输入smtp服务器验证登录名
'msg.MailServerPassword = request.Form("mima") '输入smtp服务器验证密码，,如果是在本机测试,请忽略此行代码,即无须输入密码
msg.From =request.Form("mailfrom")  '发件人Email地址 
msg.FromName = MailServerUserName '发件人姓名 
msg.AddRecipient(request.Form("mailto")) '收件人Email地址 
msg.Subject = request.Form("mailsubject") '信件主题 
msg.Body = request.Form("mailbody") '正文
msg.AddAttachment Server.MapPath(request.Form("file_path")),True  '附件
msg.Send (request.Form("smtpaddress")) 'smtp服务器地址（企业邮局地址）
msg.close() 
set msg = nothing
If err<>0 Then
  Response.Write("<script language='JavaScript'>alert('发送失败,请重新填写信息后再进行操作!');history.back();</script>")
Else
  Response.Write("<script language='JavaScript'>alert('邮件发送成功!');window.location.href='fujian.asp';</script>")
End If
%> 
