<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../Conn/conn.asp" -->
<%
if request.Form("subject")<>"" then
	subject=request.Form("subject")
	content=request.Form("content")
	person=request.Form("person")
	Ins="Insert into Tab_Placard (subject,content,person) values('"&subject&"','"&content&"','"&person&"')"
	conn.execute(Ins)%>
	<script language="javascript">
	alert("数据添加成功！");
	opener.location.reload();
	window.close();
	</script>
<%end if%>
<%
Set rs_user = Server.CreateObject("ADODB.Recordset")
sql_user = "SELECT UserName FROM Tab_User WHERE UserName = '"&session("UserName")&"'"
rs_user.Open sql_user,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>添加新公告信息！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<script language="javascript">
function Mycheck()
{
if (form1.subject.value=="")
{ alert("请输入公告主题！");form1.subject.focus();return;}
if (form1.content.value=="")
{ alert("请输入公告内容！");form1.content.focus();return;}
form1.submit();
}
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
.STYLE5 {
	font-size: 9pt;
	color: #000000;
}
-->
</style></head>

<body>
<table width="557" height="401" border="0" cellpadding="-2" cellspacing="-2"
 background="../../Images/detail0.gif">
  <tr>
    <td width="817" height="349" valign="top"><table width="100%" height="90"
	  border="0" cellpadding="-2" cellspacing="-2">
      <tr>
        <td height="43" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE4">添加新公告信息</span></td>
      </tr>
      <tr>
        <td height="45">&nbsp;</td>
      </tr>
    </table>      
      <form name="form1" method="POST" action="bbs_add.asp">
        <table width="80%" height="213"  border="0" align="center" cellpadding="-2" cellspacing="-2">
          <tr>
            <td width="13%" height="27"><div align="center" class="STYLE5">主题：</div></td>
            <td width="87%"><input name="subject" type="text" class="Style_subject" id="subject"></td>
          </tr>
          <tr>
            <td height="135" valign="top"><div align="center" class="STYLE5">内容：</div></td>
            <td><textarea name="content" cols="53" rows="8" class="Style_content"
			 id="content"></textarea></td>
          </tr>
          <tr>
            <td colspan="2"><div align="center">
                <input name="Button" type="button" class="Style_button" value="提交" onClick="Mycheck()">
                <input name="Submit2" type="reset" class="Style_button" value="重置">                
                <input name="myclose" type="button" class="Style_button" id="myclose" value="关闭"
				 onClick="javascrip:window.close()">
                <input name="person" type="hidden" id="person" value="<%=rs_user("UserName")%>">
                </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>
