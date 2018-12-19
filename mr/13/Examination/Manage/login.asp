<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute=Now()-1
Response.CacheControl="no-cache" 
%>
<!--#include file="../conn/conn.asp"-->
<%
	dim rndnum,verifycode
	Randomize
	Do While Len(rndnum)<4
	num1=CStr(Chr((57-48)*rnd+48))
	rndnum=rndnum&num1
	loop
	session("verifycode")=rndnum
	
If Not Isempty(Request.form("login")) Then
  txt_name=Request.Form("txt_name")
  txt_passwd=Request.Form("txt_passwd")
  verifycode=Request.Form("verifycode")
  verifycode2=Request.Form("verifycode2")
  If verifycode <> verifycode2 then 
	Response.write"<SCRIPT language='JavaScript'>alert('您输入的验证码不正确!');location.href='login.asp'</SCRIPT>"
	Response.End()
  Else
    Session("verifycode")=""
  End IF
  If txt_name<>"" Then
    Set rs=Server.CreateObject("ADODB.Recordset")
	sqlstr="select id,name,pwd from tb_administrator where name='"&txt_name&"'"
	rs.open sqlstr,conn,1,1
	If rs.eof Then
	  Response.Write("<script lanuage='javascript'>alert('管理员名称不正确,请核实后重新输入!');location.href='login.asp';</script>")
	Else	  
	  If rs("pwd")<>txt_passwd Then
	    Response.Write("<script lanuage='javascript'>alert('密码不正确,请确认后重新输入!');location.href='login.asp';</script>")
	  Else
	  	Session("Adminid") = rs("id")
		Session("AdminName") = txt_name
		Session("AdminPass") = txt_passwd
		Response.Redirect("adm_Main.asp")
	  End If
	End If
  Else
    errstr="请输入管理员名称!"
  End If
End If

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>管理员登录</title>
<link rel="stylesheet" href="css/css.css">
<script type="text/javascript">
<!--
function Mycheck(){
  if(document.all.txt_name.value==""){
  alert('请输入管理员名称！');return false;}
  if(document.all.txt_passwd.value==""){
  alert('请输入密码！');return false;}
  if(document.all.verifycode.value==""){
  alert('请输入验证码！');return false;}
}
-->
</script>
</head>
<body>
<table width="382" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="99" colspan="3" bgcolor="#E6F7FE">&nbsp;</td>
  </tr>
  <tr>
    <td height="80" bgcolor="#E6F7FE">&nbsp;</td>
    <td background="images/login_05.gif" bgcolor="#E6F7FE"><table width="225" height="112" border="0" align="center" cellpadding="3" cellspacing="0">
      <form name="form1" method="post" action="">
        <tr>
          <td width="56" height="22" align="right">管理员:</td>
          <td width="157" height="22"><input name="txt_name" type="text" class="textbox" id="txt_name" size="18" maxlength="50"></td>
        </tr>
        <tr>
          <td height="22" align="right">密 码:</td>
          <td height="22"><input name="txt_passwd" type="password" class="textbox" id="txt_passwd" size="19" maxlength="50"></td>
        </tr>
        <tr>
          <td height="22" align="right">验证码:</td>
          <td height="22"><input name="verifycode" type="text" class="textbox" id="textbox" size="5" maxlength="4">
              <span style="color: #FF0000"><%=session("verifycode")%></span>
              <input type="hidden" name="verifycode2" value="<%=session("verifycode")%>"></td>
        </tr>
        <tr>
          <td height="22" colspan="2" align="center"><input name="login" type="submit" id="login" value="登 录" class="button" onClick="return Mycheck()">
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="reset" name="Submit2" value="重 置" class="button"></td>
        </tr>
      </form>
    </table></td>
    <td bgcolor="#E6F7FE">&nbsp;</td>
  </tr>
  <tr>
    <td height="80" colspan="3" bgcolor="#E6F7FE">&nbsp;</td>
  </tr>
</table>
</body>
</html>