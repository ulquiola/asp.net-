<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
If Not Isempty(Request("sure")) Then 
Emails=Split(Request("email"),",") 
for i=0 to ubound(Emails)
Set msg = Server.CreateObject("JMail.Message") '不能放在循环外面
msg.silent = true 
msg.Logging = true 
msg.Charset = "gb2312" 
msg.MailServerUserName =request.Form("mailfrom")  '输入smtp服务器验证登录名,即任何一个在该smpt服务器上申请的Email地址） 
'msg.MailServerPassword = request.Form("mima") '输入smtp服务器验证密码,即Email账号对应的密码,如果是在本机测试,请忽略此行代码,即无须输入密码
msg.From =request.Form("mailfrom")  '发件人Email地址 
msg.FromName = "pyj" '发件人姓名 
msg.AddRecipient(Emails(i)) '收件人Email地址 
msg.Subject = request.Form("mailsubject") '信件主题 
msg.Body = request.Form("mailbody") '正文 
msg.Send (request.Form("smtpaddress")) 'smtp服务器地址（企业邮局地址）
msg.close()
set msg = nothing 
next
Response.Write("<script language='JavaScript'>alert('邮件群发成功!');window.location.href='qunfa.asp';</script>")
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
<title>如何实现邮件群发</title>
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
    <td height="30" colspan="4"  align="center" valign="middle">通讯录</td>
  </tr>
<tr>
  <td align="left" valign="top" id="peop">
  <table width="90%" height="153" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#99CC00">
  <tr>
    <td width="104" height="30"><div align="center">选择</div></td>
    <td width="60"><div align="center">编号</div></td>
    <td width="100"><div align="center">收件人的名字</div></td>
    <td width="309" align="center" valign="middle"><div>收件人的信箱</div></td>
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
  	<td height="30" align="center" class="style1"><input type="button" name="Submit4" value="取消选择" onClick="CheckAll(myform.email)"></td>
  	<td height="30" colspan="3" class="style1" style="line-height:18px">注意：选中复选框，可以添加邮件的收件人！单击“取消选择”按钮可以取消所有选中项，重新进行选择！</td>
  	</tr>
 </table>
 </td>
</tr>
  <tr>
  	<td height="10" colspan="4" align="center" valign="top" id="peop">
		<table width="755" border="0" align="center">
  			<tr>
    			<td width="127" height="22" align="right" valign="middle"><span>发件人：</span></td>
    			<td width="33" align="right" valign="middle">&nbsp;</td>
    			<td width="581" height="22" align="left" valign="middle"><span><input name="mailfrom" type="text" id="mailfrom" size="40"></span></td>
 			</tr>
  			<tr>
    			<td height="22" align="right" valign="middle"><span>密码：</span></td>
    			<td height="22" align="right" valign="middle">&nbsp;</td>
    			<td height="22" align="left" valign="middle"><span><input name="mima" type="password" id="mima" size="44"></span></td>
  			</tr>
  			<tr>
    			<td height="22" align="right" valign="middle"><span>SMTP服务器地址：    </span></td>
    			<td height="22" align="right" valign="middle">&nbsp;</td>
    			<td height="22" align="left" valign="middle"><span ><input name="smtpaddress" type="text" id="smtpaddress" size="40"></span></td>
  			</tr>
  			<tr>
   				 <td height="22" align="right" valign="middle"><span>标题：</span></td>
    			 <td height="22" align="right" valign="middle">&nbsp;</td>
   			    <td height="22" align="left" valign="middle"><span><input name="mailsubject" type="text" id="mailsubject" size="40"></span></td>
  			</tr>
  			<tr>
    			<td align="right" valign="middle"><span>　内容：</span></td>
   				<td align="right" valign="middle">&nbsp;</td>
   				<td align="left" valign="middle"><span ><textarea name="mailbody" cols="40" rows="5" wrap="PHYSICAL" id="mailbody"></textarea></span></td>
  			</tr>
  			<tr>
   				<td height="30" align="right" valign="middle">&nbsp;</td>
  			    <td height="30" align="center" valign="middle">&nbsp;</td>
  			    <td height="30" align="left" valign="middle"><input name="sure" type="submit" id="sure" value="发送">
  			      &nbsp;&nbsp;
		        <input type="reset" name="Submit2" value="重写"></td>
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
