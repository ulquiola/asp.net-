<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp" -->
<%
set rs_user=server.CreateObject("ADODB.Recordset")
sql="select * from Tab_User where username='"&Session("UserName")&"'"
rs_user.open sql,conn,1,3
%>
<% if trim(rs_user("purview"))<>"ϵͳ" then
%>
 <script language="javascript">
	alert("�Բ�����û�������Ա����Ȩ��!!");
	window.location.href='../main.asp';
	</script>
<% end if%>
<%
if request.Form("username")<>""then
'����û����Ƿ����
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="SELECT UserName FROM Tab_User WHERE UserName='"&Request.Form("UserName")&"'"
	rs.open sql,conn,1,3
	if rs.eof then 
	    '����Ա����Ϣ
		username=request.Form("username")
		cname=request.Form("name")
		branch=request.Form("branch")
		PWD=request.Form("PWD")
		job=request.Form("job")
		tel=request.Form("tel")
		purview=request.Form("purview")
		sex=request.Form("sex")
		email=request.Form("email")
		address=request.Form("address")
		Ins="Insert into Tab_User (username,name,branch,PWD,job,tel,purview,sex,"&_
		"email,address) values('"&username&"','"&cname&"','"&branch&"','"&PWD&"','"&_
		job&"','"&tel&"','"&purview&"','"&sex&"','"&email&"','"&address&"')"
		conn.execute(Ins) 
		%>
		<script language="javascript">
		alert("Ա����Ϣ��ӳɹ���");
	    window.location.href='personnel_add.asp';
		</script>
	<%else%>
		<script language="javascript">
		alert("��Ա����Ϣ�Ѿ����ڣ�");
		</script>
	<% 
	    end if
		end if
	%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>�����Ա��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<script language="javascript">
function Mycheck()
{
if (form1.username.value=="")
{ alert("������Ա��������");form1.username.focus();return;}
if(form1.name.value=="")
{alert("����������??");form1.name.focus();return;}
if (form1.PWD.value=="")
{ alert("���������룡");form1.PWD.focus();return;}
if (form1.tel.value=="")
{ alert("������Ա������ϵ�绰��");form1.tel.focus();return;}
if(form1.email.value=="")
{alert("������E-mail��ַ??");form1.email.focus();return;}
if(form1.email.value!="" && (form1.email.value.indexOf('@',0)==-1|| form1.email.value.indexOf('.',0)==-1))
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
.STYLE8 {
	font-size: 9pt;
	color: #000000;
}
-->
</style>
</head>
<body>
<table width="814" height="505" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="488" valign="top" background="../Images/main.gif"><table width="89%" height="84%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td>
		
		<table width="557" height="401" border="0" cellpadding="-2" cellspacing="-2" background="../images/bbs/bbs_add.gif">
  <tr>
    <td width="817" height="349" valign="top"><table width="100%" height="104"  border="0" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="6%" height="51" valign="bottom"><div align="right"><img src="../Images/add.gif" width="20" height="18"></div></td>
        <td width="94%" valign="bottom">&nbsp;<img src="../Images/yuangong.jpg" width="101" height="17"></td>
      </tr>
      <tr>
        <td height="53" colspan="2">&nbsp;</td>
      </tr>
    </table>      
      <form ACTION="personnel_add.asp" METHOD="POST" name="form1">
        <table width="100%" height="177"  border="0" align="center" cellpadding="-2" cellspacing="-2">
		  <tr>
            <td width="13%" height="27">
              <div align="center" class="STYLE8">�û�����</div></td><td width="28%"><input name="username" type="text" class="Sytle_text" id="username"></td>
            <td width="14%"><div align="center" class="STYLE8">���ţ�</div></td>
            <td><select name="branch" id="select4">
              <option value="������" selected>������</option>
              <option value="���²�">���²�</option>
              <option value="���۲�">���۲�</option>
              <option value="�ĵ���">�ĵ���</option>
            </select></td>
          </tr>
		��<tr>
            <td><div align="center" class="STYLE8">Ա��������</div></td>
            <td><input name="name" type="text" class="Sytle_text" id="name"></td>
            <td><div align="center" class="STYLE8">�Ա�</div></td>
            <td><select name="sex" id="sex">
              <option value="��" selected>��</option>
              <option value="Ů">Ů</option>
            </select></td>
            </tr>
          <tr>
            <td height="27" valign="top">
              <div align="center" class="STYLE8">���룺</div></td>
            <td><input name="PWD" type="password" class="Sytle_text" id="PWD2"></td>
            <td><div align="center" class="STYLE8">ְ��</div></td>
            <td><select name="job" id="job">
              <option value="�ܾ���" selected>�ܾ���</option>
              <option value="���ž���">���ž���</option>
              <option value="������Դ����">������Դ����</option>
              <option value="����">����</option>
              <option value="��������">��������</option>
              <option value="��ͨԱ��">��ͨԱ��</option>
              <option value="����Ա">����Ա</option>
            </select></td>
          </tr>
          <tr>
            <td height="27" valign="top">
              <div align="center" class="STYLE8">�绰��</div></td>
            <td><input name="tel" type="text" class="Sytle_text" id="tel3"></td>
            <td><div align="center" class="STYLE8">Ȩ�ޣ�</div></td>
            <td width="45%"><select name="purview" id="select">
              <option value="ϵͳ" selected>ϵͳ</option>
              <option value="��д">��д</option>
              <option value="ֻ��">ֻ��</option>
            </select></td>
          </tr>
          <tr>
            <td height="27" valign="top"><div align="center" class="STYLE8">Email��</div></td>
            <td colspan="3"><input name="email" type="text" class="Style_subject" id="email"></td>
          </tr>
          <tr>
            <td height="27" valign="top"><div align="center" class="STYLE8">��ַ��</div></td>
            <td colspan="3"><input name="address" type="text" class="Style_subject" id="address"></td>
          </tr>
          <tr>
            <td colspan="4"><div align="center">
                <input name="Button" type="button" class="Style_button" value="�ύ" onClick="Mycheck()">
                <input name="Submit2" type="reset" class="Style_button" value="����">                
                </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>

&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
