<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp" -->
<%
'��ѯԱ����Ϣ
If Request.QueryString("ID")<>""then
session("ID")=Request.QueryString("ID")
end if
Set rs_personnel = Server.CreateObject("ADODB.Recordset")
sql_P="SELECT ID,UserName,name,PWD,purview,branch,job,sex,Email,Tel,Address"&_
" FROM Tab_User WHERE ID="&session("ID")&""
rs_personnel.open sql_p,conn,0,3
%>
<%
'�޸�Ա����Ϣ
if request.Form("Name")<>"" then
	cname=request.Form("name")
	PWD=request.Form("PWD")
	purview=request.Form("purview")
	sex=request.Form("sex")
	tel=request.Form("tel")
	branch=request.Form("branch")
	job=request.Form("job")
	email=request.Form("email")
	address=request.Form("address")
	UP="Update Tab_User set name='"&cname&"',PWD='"&PWD&"',purview='"&_
	purview&"',sex='"&sex&"',tel='"&tel&"',branch='"&branch&"',job='"&job&_
	"',email='"&email&"',address='"&address&"' where ID='"&session("ID")&"'"
	conn.execute(UP)
	%>
	<script language="javascript">
	alert("�����޸ĳɹ���");
	opener.location.reload();
	window.close();
	</script>
<%end if%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>�޸�Ա����Ϣ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<script language="javascript">
function Mycheck()
{
if (form1.name.value=="")
{ alert("������Ա��������");form1.name.focus();return;}
if (form1.PWD.value=="")
{ alert("���������룡");form1.PWD.focus();return;}
if (form1.tel.value=="")
{ alert("������Ա������ϵ�绰��");form1.tel.focus();return;}
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
        <td height="102"><table width="434" height="52" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE4">&nbsp;�޸�Ա����Ϣ</span></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table></td>
      </tr>
    </table>      
      <br>
      <form ACTION="personnel_modify.asp" METHOD="POST" name="form1">
        <table width="95%" height="222"  border="0" align="center" cellpadding="-2" cellspacing="-2">
          <tr>
            <td colspan="2"><span class="style1">&nbsp;�û�����<%=(rs_personnel("UserName"))%></span></td>
            <td class="style1" align="center">���룺</td>
            <td colspan="3"><span class="style1">
              <input name="PWD" type="password" class="Sytle_text" id="name" value="<%=rs_personnel("PWD")%>">
            </span></td>
          </tr>
		  <tr>
            <td width="11%" height="27"><div align="center" class="style1">������</div></td>
            <td width="25%" height="27"><span class="style1">
              <input name="name" type="text" class="Sytle_text" id="username2" value="<%=(rs_personnel("Name"))%>" size="15">
            </span> </td>
            <td width="10%" height="27" class="style1"><div align="center">Ȩ�ޣ�</div></td>
            <td height="27"><select name="purview" id="purview">
              <option value="ϵͳ" <%If (Not isNull((rs_personnel("purview")))) Then If ("ϵͳ" = CStr((rs_personnel("purview")))) Then Response.Write("SELECTED") : Response.Write("")%>>ϵͳ</option>
              <option value="��д" <%If (Not isNull((rs_personnel("purview")))) Then If ("��д" = CStr((rs_personnel("purview")))) Then Response.Write("SELECTED") : Response.Write("")%>>��д</option>
              <option value="ֻ��" <%If (Not isNull((rs_personnel("purview")))) Then If ("ֻ��" = CStr((rs_personnel("purview")))) Then Response.Write("SELECTED") : Response.Write("")%>>ֻ��</option>
            </select></td>
            <td width="9%" height="27" class="style1"><div align="center">�Ա�</div></td>
            <td width="29%" height="27"><span class="style1">
              <select name="sex" id="select">
                <option value="��" <%If (Not isNull((rs_personnel("sex")))) Then If ("��" = CStr((rs_personnel("sex")))) Then Response.Write("SELECTED") : Response.Write("")%>>��</option>
                <option value="Ů" <%If (Not isNull((rs_personnel("sex")))) Then If ("Ů" = CStr((rs_personnel("sex")))) Then Response.Write("SELECTED") : Response.Write("")%>>Ů</option>
              </select>
</span></td>
          </tr>
          <tr>
            <td height="27"><div align="center" class="style1">�绰��</div></td>
            <td height="27"><input name="tel" type="text" class="Sytle_text" id="tel" value="<%=(rs_personnel("Tel"))%>" size="15"></td>
            <td width="10%" height="27" class="style1"><div align="center">���ţ�</div></td>
            <td width="16%" height="27"><select name="branch" id="branch">
              <option value="������" <%If (Not isNull((rs_personnel("branch")))) Then If ("������" = CStr((rs_personnel("branch")))) Then Response.Write("SELECTED") : Response.Write("")%>>������</option>
              <option value="���²�" <%If (Not isNull((rs_personnel("branch")))) Then If ("���²�" = CStr((rs_personnel("branch")))) Then Response.Write("SELECTED") : Response.Write("")%>>���²�</option>
              <option value="���۲�" <%If (Not isNull((rs_personnel("branch")))) Then If ("���۲�" = CStr((rs_personnel("branch")))) Then Response.Write("SELECTED") : Response.Write("")%>>���۲�</option>
              <option value="�ĵ���" <%if (Not isNull((rs_personnel("branch")))) then if ("�ĵ���" = CStr((rs_personnel("branch")))) then response.Write("SELECTED") :response.Write("")%>>�ĵ���</option>
            </select></td>
            <td width="9%" height="27"><div align="center"><span class="style1">ְ��</span></div></td>
            <td height="27"><span class="style1">              <select name="job" id="job">
                <option value="�ܾ���" <%If (Not isNull((rs_personnel("job")))) Then If ("�ܾ���" = CStr((rs_personnel("job")))) Then Response.Write("SELECTED") : Response.Write("")%>>�ܾ���</option>
                <option value="���ž���" <%If (Not isNull((rs_personnel("job")))) Then If ("���ž���" = CStr((rs_personnel("job")))) Then Response.Write("SELECTED") : Response.Write("")%>>���ž���</option>
                <option value="������Դ����" <%If (Not isNull((rs_personnel("job")))) Then If ("������Դ����" = CStr((rs_personnel("job")))) Then Response.Write("SELECTED") : Response.Write("")%>>������Դ����</option>
                <option value="����" <%If (Not isNull((rs_personnel("job")))) Then If ("����" = CStr((rs_personnel("job")))) Then Response.Write("SELECTED") : Response.Write("")%>>����</option>
                <option value="��������" <%If (Not isNull((rs_personnel("job")))) Then If ("��������" = CStr((rs_personnel("job")))) Then Response.Write("SELECTED") : Response.Write("")%>>��������</option>
                <option value="��ͨԱ��" <%If (Not isNull((rs_personnel("job")))) Then If ("��ͨԱ��" = CStr((rs_personnel("job")))) Then Response.Write("SELECTED") : Response.Write("")%>>��ͨԱ��</option>
                <option value="����Ա" <%if (Not isNull((rs_personnel("job")))) then if ("����Ա") = cstr((rs_personnel("job"))) then response.Write("SELECTED") : response.Write("")%>>����Ա</option>
              </select>
</span></td>
          </tr>
          <tr>
            <td height="27" class="style1"><div align="center">Email��</div></td>
            <td height="27" colspan="5"><input name="email" type="text" class="Style_subject" id="email" value="<%=(rs_personnel("Email"))%>" size="15"></td>
          </tr>
          <tr>
            <td height="27" class="style1"><div align="center">��ַ��</div></td>
            <td height="27" colspan="5"><input name="address" type="text"
			 class="Style_subject" id="address" value="<%=(rs_personnel("Address"))%>"></td>
          </tr>
          <tr>
            <td colspan="6"><div align="center">
                <br>
                <input name="Button" type="button" class="Style_button" value="�޸�" onClick="Mycheck()">
                <input name="Submit2" type="reset" class="Style_button" value="����">
                </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>
