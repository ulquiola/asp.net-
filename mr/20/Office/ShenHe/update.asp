<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../conn/conn.asp"-->
<%
if request.Querystring("ID")<>"" then
	session("id")=request.Querystring("ID")
end if
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Tab_shehe where ID='"&session("id")&"'"
rs.open sql,conn,1,3
if request.Form("title")<>"" then
	title=request.Form("title")
	content=request.Form("content")
	Time1=date()
	UP="Update Tab_shehe set title='"&title&"',content='"&content&"',time1='"&date()&"' where ID='"&session("id")&"'"
	conn.execute(UP)
%>
<script language="javascript">
alert("信息修改成功!");
opener.location.reload();
window.close();
</script>
<%end if%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>修改申请</title>
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
.STYLE3 {	font-size: 9pt;
	color: #000000;
}
-->
</style></head>
<script language="javascript">
function Mycheck()
{
if(form1.title.value=="")
{alert("请输入申请标题!");form1.title.focus();return;}
if(form1.content.value=="")
{alert("请输入申请内容!");form1.content.focus();return;}
form1.submit();
}
</script>
<body>
<form name="form1" method="post" action="update.asp">
<table width="520" height="465" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="495" height="465" valign="top" background="../Images/tongxun_add.gif"><table width="500" height="458" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="27" valign="bottom">&nbsp;&nbsp;<span class="STYLE2">&nbsp;&nbsp;修改申请</span></td>
      </tr>
      <tr>
        <td height="431" valign="middle"><table width="427" height="298" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="20"><div align="center" class="STYLE3">申请主题</div></td>
            <td width="283"><div align="center" class="STYLE3">
                <div align="left">
                  <input name="title" type="text" id="title" value="<%=rs("title")%>" size="20" />
                </div>
            </div></td>
          </tr>
          <tr>
            <td width="72" height="187" valign="middle" class="STYLE3">&nbsp;&nbsp;
                <div align="center">申请内容<br />
              </div></td>
            <td height="187" colspan="3" valign="middle" class="STYLE3"><textarea name="content" cols="40" rows="17" id="content"><%=rs("content")%></textarea></td>
          </tr>
          <tr>
            <td height="15">&nbsp;</td>
            <td width="70" height="15">&nbsp;</td>
            <td height="15" colspan="2">&nbsp;</td>
          </tr>
        </table>
          <table width="237" height="28" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td><a href="#"><img src="../Images/10.GIF" width="80" height="31" border="0" onClick="Mycheck();" /></a><a href="#"><img src="../Images/waichuan8.GIF" width="62" height="31" border="0" onClick="JScript:window.close();" /></a><a href="#"><img src="../Images/qiyebutton1.gif" width="81" height="31" border="0" onClick="JScript:form1.reset();" /></a></td>
            </tr>
          </table>          <p>&nbsp;</p></td>
      </tr>
    </table></td>
  </tr>
</table>
</form>
</body>
</html>
