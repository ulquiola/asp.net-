<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp" -->
<%
Set rs_user = Server.CreateObject("ADODB.Recordset")
sql_User = "SELECT UserName FROM Tab_User WHERE UserName = '"&session("UserName")&"'"
rs_user.open sql_User,conn,1,3
'查询收件人列表
set rs_lperson = Server.CreateObject("ADODB.Recordset")
sql_lp="SELECT UserName,purview FROM Tab_User WHERE  UserName<>'"&Session("UserName")&"'  ORDER BY UserName ASC"
rs_lperson.open sql_lP,conn,1,3
If request.Form("Lperson")<>"" then
	'保存发文信息
	Lperson=Request.Form("Lperson")
	subject=Request.Form("subject")
	content=Request.Form("content")
	Sperson=Request.Form("Sperson")
	Ins="Insert into Tab_Send(Lperson,subject,content,Sperson) values('"&Lperson&_
	"','"&subject&"','"&content&"','"&Sperson&"')"
	conn.execute(Ins)%>
	<script language="javascript">
	alert("发文成功！");
	opener.location.reload();
	window.close();
	</script>
<%end if%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>发送公文！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<script language="javascript">
function Mycheck()
{
if (form1.Lperson.value=="")
{ alert("请输入收文人姓名！");form1.Lperson.focus();return;}
if (form1.subject.value=="")
{ alert("请输入公文主题！");form1.subject.focus();return;}
if (form1.content.value=="")
{ alert("请输入公文内容！");form1.content.focus();return;}
form1.submit();}
</script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
.style1 {color: #C60001}
.style2 {color: #669999}
.STYLE4 {
	font-size: 9pt;
	color: #FFFFFF;
}
.STYLE5 {font-size: 9pt}
.STYLE6 {font-size: 9pt; color: #C60001; }
-->
</style></head>

<body>
<table width="460" height="403" border="0" cellpadding="-2" cellspacing="-2" background="../Images/waichudeng.gif">
  <tr>
    <td width="817" height="349" valign="top"><table width="100%" height="102"  border="0" cellpadding="-2" cellspacing="-2">
      <tr>
        <td height="45" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE4">&nbsp;发送公文</span></td>
      </tr>
      <tr>
        <td height="57">&nbsp;</td>
      </tr>
    </table>      
      <form ACTION="send_add.asp" METHOD="POST" name="form1">
        <table width="93%" height="207"  border="0" align="center" cellpadding="-2" cellspacing="-2">
          <tr>
            <td width="15%" height="27"><div align="left" class="style1">
              <div align="center" class="STYLE5">收文人：</div>
            </div></td>
            <td width="85%"><select name="Lperson" id="Lperson">
              <%
While (NOT rs_lperson.EOF)
%>
              <option value="<%=rs_lperson("UserName")%>"><%=rs_lperson("UserName")%></option>
              <%
  rs_lperson.MoveNext()
Wend
If (rs_lperson.CursorType > 0) Then
  rs_lperson.MoveFirst
Else
  rs_lperson.Requery
End If
%>
            </select></td>
          </tr>
          <tr>
            <td height="27"><div align="center"><span class="STYLE6">主题：</span></div></td>
            <td width="85%"><input name="subject" type="text" class="Style_subject" id="subject"></td>
          </tr>
          <tr>
            <td height="10" valign="top"><div align="center" class="style1">
              <div align="center" class="STYLE5">内容：</div>
            </div></td>
            <td><textarea name="content" cols="45" rows="6" class="Style_content" id="content"></textarea></td>
          </tr>
          <tr>
            <td colspan="2"><div align="center">
                <input name="Button" type="button" class="Style_button" value="提交" onClick="Mycheck()">
                <input name="Submit2" type="reset" class="Style_button" value="重置">                
                <input name="myclose" type="button" class="Style_button" id="myclose" value="关闭"
				 onClick="javascrip:window.close()">
                <input name="Sperson" type="hidden" id="Sperson" value="<%=rs_user("UserName")%>">
                </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>
