<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp" -->
<%
'�޸ĸ�����Ϣ������
if request.Form("userName")<>"" then
	cName=request.Form("userName")
	sex=request.Form("sex")
	PWD=request.Form("NewPWD")
	tel=request.Form("tel")
	email=request.Form("email")
	address=request.Form("address")
	UP="Update Tab_User set name='"&cName&"',sex='"&sex&"',PWD='"&PWD&"',tel='"&tel&_
	"',email='"&email&"',address='"&address&"' where ID='"&session("ID")&"'"
	conn.execute(UP)%>
	<script language="javascript">
	alert("���������޸ĳɹ���");
	opener.parent.location.reload();
	window.close();
	</script>
<%end if%>
<%
If Request.QueryString("ID")<>""then
session("ID")=Request.QueryString("ID")
end if
Set rs_personnel = Server.CreateObject("ADODB.Recordset")
sql_P= "SELECT ID,UserName,Name,PWD,purview,branch,sex,Email,Tel,Address FROM Tab_User WHERE ID="&_
Session("ID")&""
rs_personnel.open sql_P,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>�޸ĸ�����Ϣ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<script language="javascript">
function Mycheck()
{
if (form1.username.value=="")
{ alert("������Ա��������");form1.username.focus();return;}
if (form1.PWD.value=="")
{ alert("������ԭ���룡");form1.PWD.focus();return;}
if (form1.PWD.value!=form1.hPWD.value)
{ alert("������ԭ���벻��ȷ�����������룡");form1.PWD.value="";
form1.PWD.focus();return;}
if (form1.NewPWD.value=="")
{ alert("�����������룡");form1.NewPWD.focus();return;}
if (form1.PWDOK.value=="")
{ alert("ȷ�������룡");form1.PWDOK.focus();return;}
if (form1.NewPWD.value!=form1.PWDOK.value)
{alert("����������������벻һ�£����������룡");form1.NewPWD.value="";
form1.PWDOK.value="";form1.NewPWD.focus();return;}
if (form1.tel.value=="")
{ alert("������Ա������ϵ�绰��");form1.tel.focus();return;}
if(form1.email.value!="" && (form1.email.value.indexOf('@',0)==-1||
 form1.email.value.indexOf('.',0)==-1))
{alert("�������E-mail��ַ���ԣ�");form1.email.focus();return;}
if (form1.address.value=="")
{ alert("������Ա����סַ��");form1.address.focus();return;}
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
-->
</style></head>

<body>
<table width="460" height="403" border="0" cellpadding="-2" cellspacing="-2" background="../Images/waichudeng.gif">
  <tr>
    <td width="817" height="349" valign="top"><table width="100%" height="102"  border="0" cellpadding="-2" cellspacing="-2">
      <tr>
        <td height="102"><table width="424" height="47" border="0">
          <tr>
            <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE4">�޸ĸ�����Ϣ</span></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table></td>
      </tr>
    </table>      
      <form ACTION="personnel_modifyPWD.asp" METHOD="POST" name="form1">
        <table width="95%" height="177"  border="0" align="center" cellpadding="-2" cellspacing="-2">
          <tr>
            <td colspan="4" class="style1">&nbsp;&nbsp;�û�����<%=rs_personnel("UserName")%></td>
          </tr>
		  <tr>
            <td width="14%" height="27"><div align="center" class="style1">������</div></td>
            <td width="31%"><input name="username" type="text" class="Sytle_text" id="username" value="<%=rs_personnel("Name")%>" size="18"></td>
            <td width="15%" class="style1">ԭ���룺</td>
            <td width="40%"><input name="PWD" type="password" class="Sytle_text" id="PWD">
            <input name="hPWD" type="hidden" id="hPWD" value="<%=(rs_personnel("PWD"))%>"></td>
          </tr>
          <tr>
            <td height="27" valign="top"><div align="center" class="style1">�Ա�</div></td>
            <td><span class="style1">
              <select name="sex" id="select">
                <option value="��" <%If (Not isNull(rs_personnel("sex"))) Then If ("��" = CStr(rs_personnel("sex"))) Then Response.Write("SELECTED") : Response.Write("")%>>��</option>
                <option value="Ů" <%If (Not isNull(rs_personnel("sex"))) Then If ("Ů" = CStr(rs_personnel("sex"))) Then Response.Write("SELECTED") : Response.Write("")%>>Ů</option>
              </select>
            </span> </td>
            <td class="style1">�����룺</td>
            <td><input name="NewPWD" type="password" class="Sytle_text" id="NewPWD"></td>
          </tr>
          <tr>
            <td height="27" valign="top"><div align="center" class="style1">�绰��</div></td>
            <td><input name="tel" type="text" class="Sytle_text" id="tel" value="<%=(rs_personnel("Tel"))%>" size="18"></td>
            <td class="style1">ȷ�����룺</td>
            <td><input name="PWDOK" type="password" class="Sytle_text" id="PWDOK"></td>
          </tr>
          <tr>
            <td height="27" valign="top" class="style1"><div align="center">Email��</div></td>
            <td colspan="3"><input name="email" type="text" class="Style_subject" id="email" value="<%=(rs_personnel("Email"))%>"></td>
          </tr>
          <tr>
            <td height="27" valign="top" class="style1"><div align="center">��ַ��</div></td>
            <td colspan="3"><input name="address" type="text" class="Style_subject" id="address" value="<%=(rs_personnel("Address"))%>"></td>
          </tr>
          <tr>
            <td colspan="4"><div align="center">
                <input name="Button" type="button" class="Style_button" value="�޸�" onClick="Mycheck()">
                <input name="Submit2" type="reset" class="Style_button" value="����">
                <input name="myclose" type="button" class="Style_button" id="myclose" value="�ر�" onClick="javascrip:window.close()">
            </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>
<%rs_personnel.Close()%>
