<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn/conn.asp"-->
<%
'修改记录
Sub edit()
  str1=Request.Form("txt_name2")
  str2=Request.Form("txt_passwd2")
  id=Request.Form("id")  
  sqlstr="update tb_user set usename='"&str1&"',pwd='"&str2&"' where id="&id&""
  conn.Execute(sqlstr)  
  Response.Redirect("adm_manage.asp")
End Sub
'执行子过程
If Not Isempty(Request("edit")) Then call edit()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>密码修改</title>
<link rel="stylesheet" href="css/css.css">
<style type="text/css">
<!--
body {
	margin-top: 0px;
}
-->
</style></head>
<body>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr valign="top">
    <td height="5" background="images/mid_beijing.jpg"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center" valign="top">
            <table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td height="22" class="black_12">管理员资料修改</td>
              </tr>
            </table>            <table width="500" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#E3E3E3">
              <tr align="center">
                <td width="169" height="22" bgcolor="#FFFFFF" class="black_12">管理员名称</td>
          <td width="194" height="22" bgcolor="#FFFFFF" class="black_12">密 码</td>
          <td width="137" height="22" bgcolor="#FFFFFF" class="black_12">操 作</td>
        </tr>
              <%
Set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select id,usename,pwd from tb_user"
rs.open sqlstr,conn,1,1
while not rs.eof
%>
              <form name="form2<%=rs("id")%>" method="post" action="adm_manage.asp">
                <tr align="center">
                  <td height="22" bgcolor="#FFFFFF" class="black_12"><input name="txt_name2" value="<%=rs("usename")%>" type="text" class="textbox" id="txt_name22" size="18" maxlength="50"></td>
            <td height="22" bgcolor="#FFFFFF" class="black_12"><input name="txt_passwd2" value="<%=rs("pwd")%>" type="text" class="textbox" id="txt_passwd2" size="18" maxlength="50"></td>
            <td height="22" bgcolor="#FFFFFF" class="black_12"><input name="id" type="hidden" id="id" value="<%=rs("id")%>">
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
                  </table></td>
        </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
