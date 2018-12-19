<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
If Not Isempty(Request("send")) Then 
on error resume next
Set msg = Server.CreateObject("JMail.Message") 
msg.silent = true 
msg.Logging = true 
'msg.Charset = "gb2312" 
msg.MailServerUserName =request.Form("mailfrom")  '输入smtp服务器验证登录名,即任何一个在该smpt服务器上申请的Email地址） 
'msg.MailServerPassword = request.Form("mima")	  '输入smtp服务器验证密码,即Email账号对应的密码,如果是在本机测试,请忽略此行代码,即无须输入密码
msg.From =request.Form("mailfrom")  '发件人Email地址 
msg.FromName = MailServerUserName  '发件人姓名 
msg.AddRecipient(request.Form("emails")) '收件人Email地址 
msg.Subject = request.Form("mailsubject") '信件主题 
msg.Body = request.Form("mailbody") '正文 
msg.Send (request.Form("smtpaddress")) 'smtp服务器地址（企业邮局地址）
msg.close() 
set msg = nothing 
If err<>0 Then
  Response.Write("<script language='JavaScript'>alert('邮件发送失败,请核实输入内容是否准确!');history.back();</script>")
Else
  Response.Write("<script language='JavaScript'>alert('使用Jmail组件发送邮件成功!');window.location.href='jmail.asp';</script>")
End If
End If
%> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>如何使用Jmail组件发送邮件</title>
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
	  alert("请输入完整信息内容!");return false;}		
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
        <td width="17%" align="right" valign="middle">收件人：</td>
        <td width="5%" align="left" valign="middle">&nbsp;</td>
        <td width="78%" align="left" valign="middle">
          <input name="emails" type="text" id="emails" size="40" value="<%=Request.QueryString("mailto")%>">      
          <span class="style1">*</span></td>
      </tr>
      <tr>
        <td align="right" valign="middle">发件人：</td>
        <td align="left" valign="middle">&nbsp;</td>
        <td align="left" valign="middle"><span class="style4">
      <input name="mailfrom" type="text" id="mailfrom" size="40">
      <span class="style1">*</span> </span></td>
      </tr>
      <tr>
        <td align="right" valign="middle">密码：</td>
        <td align="left" valign="middle">&nbsp;</td>
        <td align="left" valign="middle"><span class="style4">
      <input name="mima" type="password" id="mima" size="44">
      <span class="style1">*</span> </span></td>
      </tr>
      <tr>
        <td align="right" valign="middle">SMTP服务器地址：</td>
        <td align="left" valign="middle">&nbsp;</td>
        <td align="left" valign="middle"><span class="style4">
      <input name="smtpaddress" type="text" id="smtpaddress" size="40">
    (如:smtp.163.com等)<span class="style1">*</span></span></td>
      </tr>
      <tr>
        <td align="right" valign="middle">标题：</td>
        <td align="left" valign="middle">&nbsp;</td>
        <td align="left" valign="middle"><span class="style4">
      <input name="mailsubject" type="text" id="mailsubject" size="40">
      <span class="style1">*</span> </span></td>
      </tr>
      <tr>
        <td align="right" valign="middle">内容：</td>
        <td align="left" valign="middle"><span>
        </span></td>
        <td align="left" valign="middle"><span class="style4">
      <textarea name="mailbody" cols="43" rows="6" wrap="PHYSICAL" id="mailbody"></textarea>
      <span class="style1">*</span>    </span></td>
      </tr>
      <tr>
        <td align="right" valign="middle">&nbsp;</td>
        <td align="center" valign="middle">&nbsp;</td>
        <td align="left" valign="middle"><input name="send" type="submit" id="send" value="发送" onClick="return Mycheck(this.form)">
　
  <input type="reset" name="Submit2" value="重写"></td>
      </tr>
      <tr>
        <td colspan="3" align="center" valign="middle"><span>*以上内容必须填写</span></td>
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
