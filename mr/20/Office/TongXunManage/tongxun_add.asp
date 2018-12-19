<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../conn/conn.asp"-->
<%
if request.Form("name1")<>"" then
name1=request.Form("name1")
session("name1")=request.Form("name1")
sql="insert into Tab_tongxun(name1) values('"&name1&"')"
conn.execute(sql)%>
<script language="javascript">
alert("通讯组类型添加成功!!")
opener.location.reload();
window.close();
</script>
<%end if%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>添加新通讯组类型</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
.STYLE2 {
	font-size: 10pt;
	color: #FFFFFF;
}
.STYLE4 {
	font-size: 10pt;
	color: #000000;
}
-->
</style></head>
<script language="javascript">
function Mycheck()
{
if(form1.name1.value=="")
{alert("请输入通讯组名称!!");form1.name1.focus();return;}
form1.submit();
}
</script>
<body>
<form name="form1" method="post" action="tongxun_add.asp">
<table width="302" height="230" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="274" height="228" valign="top" background="../Images/tongxun_add.gif"><table width="289" height="225" border="0" align="right" cellpadding="0" cellspacing="0">
      <tr>
        <td height="27" valign="bottom">&nbsp;&nbsp;<span class="STYLE2">添加新通讯组类型 </span></td>
      </tr>
      <tr>
        <td height="198"><table width="282" height="92" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="83" height="46"><span class="STYLE4">通讯组名称</span></td>
            <td width="199"><input name="name1" type="text" id="name1" size="25" /></td>
          </tr>
          <tr>
            <td height="46" colspan="2"><table width="237" height="28" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td><a href="#"><img src="../Images/10.GIF" width="80" height="31" border="0" onClick="Mycheck();"></a><a href="#"><img src="../Images/waichuan8.GIF" width="62" height="31" border="0" onClick="JScript:window.close();"></a><a href="#"><img src="../Images/qiyebutton1.gif" width="81" height="31" border="0" onClick="JScript:form1.reset();"></a></td>
                </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>
</form>
</body>
</html>
