<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute=Now()-1
Response.CacheControl="no-cache" 
%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<%
session("verifycode")=randStr(4)
If Not Isempty(Request("login")) Then
  txt_name=Str_filter(Request.Form("txt_name"))
  txt_passwd=Str_filter(Request.Form("txt_passwd"))
  verifycode=Str_filter(Request.Form("verifycode"))
  verifycode2=Str_filter(Request.Form("verifycode2"))
  If verifycode <> verifycode2 then 
	Response.write"<SCRIPT language='JavaScript'>alert('���������֤�벻��ȷ!');location.href='login.asp'</SCRIPT>"
	Response.End()
  Else
    Session("verifycode")=""
  End IF
  If txt_name<>"" Then
    Set rs=Server.CreateObject("ADODB.Recordset")
	sqlstr="select Mname,Mpasswd from tab_manager where Mname='"&txt_name&"'"
	rs.open sqlstr,conn,1,1
	If rs.eof Then
	  Response.Write("<script lanuage='javascript'>alert('����Ա���Ʋ���ȷ,���ʵ����������!');location.href='login.asp';</script>")
	Else	  
	  If rs("Mpasswd")<>txt_passwd Then
	    Response.Write("<script lanuage='javascript'>alert('���벻��ȷ,��ȷ�Ϻ���������!');location.href='login.asp';</script>")
	  Else
	  	Session("Mname")=rs("Mname")		
		Response.Redirect("index.asp")
	  End If
	End If
  Else
    errstr="���������Ա����!"
  End If
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����Ա��¼</title>
<link rel="stylesheet" href="css/css.css">
<script type="text/javascript">
<!--
function Mycheck(){
  if(document.all.txt_name.value==""){
  alert('���������Ա���ƣ�');return false;}
  if(document.all.txt_passwd.value==""){
  alert('���������룡');return false;}
  if(document.all.verifycode.value==""){
  alert('��������֤�룡');return false;}
}
-->
</script>
</head>
<body>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<table width="360" height="245" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="160" background="images/login_01.jpg"><table width="300" border="0" align="center" cellpadding="3" cellspacing="0">
      <form name="form1" method="post" action="">
        <tr>
          <td height="22" colspan="2" align="center">&nbsp;</td>
        </tr>
        <tr>
          <td height="22" align="right">����Ա:</td>
          <td height="22"><input name="txt_name" type="text" class="textbox" id="txt_name" size="18" maxlength="50"></td>
        </tr>
        <tr>
          <td height="22" align="right">�� ��:</td>
          <td height="22"><input name="txt_passwd" type="password" class="textbox" id="txt_passwd" size="19" maxlength="50"></td>
        </tr>
        <tr>
          <td height="22" align="right">��֤��:</td>
          <td height="22"><input name="verifycode" id="verifycode" class="textbox" onFocus="this.select(); " onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" size="6" maxlength="4">
              <span style="color: #FF0000"><%=session("verifycode")%></span>
              <input type="hidden" name="verifycode2" value="<%=session("verifycode")%>"></td>
        </tr>
        <tr>
          <td height="22" colspan="2" align="center"><input type="submit" name="login" value="��¼" onClick="return Mycheck();">            
            &nbsp;&nbsp;
            <input type="reset" name="Submit2" value="����"></td></tr>
      </form>
    </table></td>
  </tr>
</table>
</body>
</html>
