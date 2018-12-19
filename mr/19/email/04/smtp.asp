<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet">
<%
getlogname = "mr"
getlogaddr =getlogname &"@mr.com"
getlogname = replace(getlogname,"'"," ")
getlogaddr = replace(getlogaddr,"'"," ")
if(getlogname<> "") then
	Set objsession = server.CreateObject("CDONTS.Session")
	objsession.LogonSMTP getlogname,getlogaddr
	Set objinbox = objsession.Inbox
	Set objmessages = objinbox.messages
end if
%>
 
<br>
<table width="515" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#98C2EC">
  <tr>
    <td height="30" background="Images/smtp.jpg" style="text-indent:45px">&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><table width="497" height="103" cellspacing="0" cellpadding="0"
		    border="1" bordercolor="#FFFFFF" bordercolordark="#3B6C9D"
			 bordercolorlight="#FFFFFF" align="center"> 
              <tr> 
                <td height="30" colspan=4 ><table width="100%" height="21" 
				border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="78%" height="21">&nbsp;您的邮件中共有[<%=objmessages.count%>] 个邮件 
					</td>
                    <td width="22%"><font style="font-size:9pt">收件人:[<%=getlogname%>]</font></td>
                  </tr>
                </table></td> 
              </tr> 
              <tr align="center"> 
                <td width="153" height="20">发件人</td> 
                <td width="192" height="20">主题</td> 
                <td width="144" height="20">发件时间</td> 
              </tr> 
				<%
				For i=1 To objmessages.count
					Set objmessage = objmessages.item(i)
				%> 
              <tr align="center"> 
                <td height="20">&lt;<%=objmessage.sender%>&gt;</td> 
                <td height="20"><a href="smtp.asp?id=<%=i%>&getlogname=<%=getlogname%>&getlogaddr=<%=getlogaddr%>">
				<%=objmessage.subject%></a></td> 
                <td height="20"><%=objmessage.timesent%></td> 
              </tr> 
				<%
				Next
				objsession.Logoff
				%>
    </table><BR>
<% 
id=Request.QueryString("id")
If id<>"" Then
  getlogname = Request.QueryString("getlogname")
  getlogaddr = Request.QueryString("getlogaddr")
  getid = replace(id,"'","''")
if(getid <> "") then
	Set objsession = server.createobject("CDONTS.Session")
	objsession.LogonSMTP getlogname,getlogaddr
	Set  objinbox = objsession.Inbox
	Set objmessages = objinbox.messages
	Set objmessage = objmessages.item(getid)
end if
%>
	<table width=499 align=center cellpadding=0 cellspacing=0 border=1 bordercolordark="#ffffff" bordercolorlight="#C8D6E3"> 
              <tr align=center> 
                <td height="26" bgcolor="#C8D6E3">发件人:</td> 
                <td height="26" colspan=3 align="left">&lt;<%=objmessage.sender%>&gt;</td> 
              </tr> 
              <tr align=center> 
                <td height="26" bgcolor="#C8D6E3">主&nbsp;&nbsp;题:</td> 
                <td width="28%" height="26" align="left"><%=objmessage.subject%></td> 
                <td width="18%" height="26" bgcolor="#C8D6E3">发件时间</td> 
                <td width="34%" height="26" align="left"><%=objmessage.timesent%></td> 
              </tr> 
              <tr align=center > 
                <td width="20%" height="26" bgcolor="#C8D6E3">收件人:</td> 
                <td height="26" colspan=3 align="left">&lt;<%=getlogaddr%>&gt;</td> 
              </tr> 
              <tr align=center bgcolor="#C8D6E3"> 
                <td height="26" colspan=4 valign=middle><div align="left"> &nbsp;&nbsp;&nbsp;内&nbsp; 容:</div></td> 
              </tr> 
              <tr> 
                <td height="30" colspan=4 align="left" valign=top> <%
content = objmessage.text
content = server.HTMLEncode(content)
content = replace(content,chr(13)&chr(10),"<br>")
content = replace(content,chr(32),"&nbsp;")
response.write(content)
%></td> 
              </tr>               
      </table>
	<%End If%>	</td>
  </tr>
</table>
<br>
<table width="776">
  <tr>
    <td width="776" height="90"><img src="images/bottom.jpg" width="776" height="90"></td>
  </tr>
</table>
