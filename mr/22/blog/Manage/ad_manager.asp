<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="checklogin.asp"-->
<%
'修改记录
Sub edit()
  str1=Str_filter(Request.Form("txt_name2"))
  str2=Str_filter(Request.Form("txt_passwd2"))
  id=Request.Form("id")  
  sqlstr="update tab_manager set Mname='"&str1&"',Mpasswd='"&str2&"' where id="&id&""
  conn.Execute(sqlstr)  
  Response.Redirect("ad_manager.asp")
End Sub
'执行子过程
If Not Isempty(Request("edit")) Then call edit()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>博客</title>
<link rel="stylesheet" href="css/css.css">
</head>
<body>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><!--#include file="top.asp"--></td>
  </tr>
  <tr valign="top">
    <td height="5" background="images/mid_beijing.jpg"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="29%" align="center" valign="top"><!--#include file="left.asp"--></td>
          <td width="71%" valign="top"><table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td height="22"><img src="images/dian.gif" width="7" height="7"> 管理员资料修改</td>
              </tr>
              <tr>
                <td height="1"><img src="images/xian.gif" width="366" height="1"></td>
              </tr>
            </table><table width="500" border="0" align="center" cellpadding="0" cellspacing="2">
      <tr align="center">
        <td width="169" height="22">管理员名称</td>
        <td width="194" height="22">密 码</td>
        <td width="137" height="22">操 作</td>
      </tr>
      <%
Set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select id,Mname,Mpasswd from tab_manager"
rs.open sqlstr,conn,1,1
while not rs.eof
%>
      <form name="form2<%=rs("id")%>" method="post" action="">
        <tr align="center">
          <td height="22"><input name="txt_name2" value="<%=rs("Mname")%>" type="text" class="textbox" id="txt_name22" size="18" maxlength="50"></td>
          <td height="22"><input name="txt_passwd2" value="<%=rs("Mpasswd")%>" type="text" class="textbox" id="txt_passwd2" size="18" maxlength="50"></td>
          <td height="22"><input name="id" type="hidden" id="id" value="<%=rs("id")%>">
              <input name="edit" type="submit" class="button" id="edit" value="修 改">
&nbsp;</td>
        </tr>
      </form>
      <%
rs.movenext
wend
rs.close
set rs=nothing
%>
    </table>
          </td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><!--#include file="bottom.asp"--></td>
  </tr>
</table>
</body>
</html>
