<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet">
<%
getaccept = trim(request.Form("accept"))
getsend = trim(request.Form("send"))
getsubject = trim(request.Form("subject"))
getcontent = trim(request.Form("content"))
if (getaccept <> "" and getsend <> "" and getsubject <> "" and getcontent <> "") then
	sendemail getsend,getaccept,getsubject,getcontent 
    Response.Write("<script language='JavaScript'>alert('使用CDONTS组件发送邮件成功!');window.location.href='send.asp';</script>")
end if
sub sendemail(frwho,towho,subject,getcontent)
	set getemail = server.CreateObject("cdonts.newmail")
	getemail.from = frwho
	getemail.to = towho
	getemail.subject = subject
	getemail.body = getcontent
	'getemail.AttachFile getattac
	getemail.importance=0	
	getemail.send
	set getemail = nothing 
end sub
%>
<script type="text/javascript">
function Mycheck(form){
  for(i=0;i<form.length;i++){
    if(form.elements[i].value==""){
	  alert("请填写完整的邮件信息!");return false;}		
  }
}
</script>
<br>
<table width="412" border="0" align="center" cellpadding="0" cellspacing="0" background="images/bg.jpg">
  <tr>
    <td height="41" background="Images/top.jpg" style="text-indent:40px">&nbsp;</td>
  </tr>
  <tr>
    <td valign="top" align="center"><form name="sendemail" action="send.asp" method="post">
              <table width=92% height="218" align=center cellpadding="0" cellspacing="0"
			   border = "0">
                <tr align=center bgcolor="#FFFFFF" >
                  <td width="21%" height="21">收件人：</td>
                  <td width="79%" align="left"><input name="accept" type="text" class="txt_grey">                  </td>
                </tr>
                <tr align=center bgcolor="#FFFFFF">
                  <td height="24">发件人：</td>
                  <td align="left"><input name="send" type="text" class="txt_grey"
				  ></td>
                </tr>
                <tr align=center bgcolor="#FFFFFF">
                  <td height="24">主 &nbsp;题：</td>
                  <td align="left"><input name="subject" type="text" class="txt_grey"
				   style="width:240"></td>
                </tr>
                <tr align=center bgcolor="#FFFFFF">
                  <td height="95" valign=top>内&nbsp;&nbsp;容：</td>
                  <td align="left" valign="top"><textarea name="content" cols="37" rows="6" 
				  class="wenbenkuang"></textarea></td>
                </tr>
                <tr align=center bgcolor="#FFFFFF">
                  <td height="27" colspan=2>                    
                    <input name="submit" type="submit" 
				  class="btn_grey" value="发送" onClick="return Mycheck(this.form)">
&nbsp;                    <input name="submit" type="reset" class="btn_grey" value="重置">
&nbsp;                  </td>
                </tr>
        </table>
          </form></td>
  </tr>
</table>
<br>
<table width="776">
	<tr>
		<td width="776" height="90"><img src="images/bottom.jpg" width="776" height="90"></td>
	</tr>
</table>