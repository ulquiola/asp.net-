<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<%
if request("ID")<>"" then
session("ID")=request("ID")
end if
'��ѯҪ�޸�����Ĺ���Ա����Ϣ
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM DB_manager WHERE ID = " &session("ID")& ""
rs.open sql,conn,1,3
%>
<%
'�޸�����
if request.Form("PWD_New1")<>"" then
	PWD_New=request.Form("PWD_New1")
	Up="Update DB_manager set pwd='"& PWD_New&"' WHERE ID = " &session("ID")& ""
	conn.execute(UP)%>
	<script language="javascript">
	alert("�����޸ĳɹ���");
	window.close();
	</script>
<%end if%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>����Ա��¼--�޸Ĺ���Ա��Ϣ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image:  url(../Images/bg.gif);
}
.style1 {color: #FFFFFF}
.style2 {color: #a2bcc5}
-->
</style></head>

<body>
<script language="javascript">
function mycheck(){
if(form1.PWD_S.value=="" && form1.manager.value!="ϵͳ")
{alert("������ԭ���룡");form1.PWD_S.focus();return;}
if(form1.PWD_S.value != form1.H_PWD.value && form1.manager.value!="ϵͳ")
{alert("�������ԭ�������");form1.PWD_S.focus();return;}
if(form1.PWD_New1.value=="")
{alert("�����������룡");form1.PWD_New1.focus();return;}
if(form1.PWD_New2.value=="")
{alert("��ȷ�������룡");form1.PWD_New2.focus();return;}
if(form1.PWD_New1.value != form1.PWD_New2.value)
{alert("����������������벻һ�£�");
form1.PWD_New1.value="";form1.PWD_New2.value="";
form1.PWD_New1.focus();return;}
form1.submit();
}
</script>
<table width="400" height="242" border="0" cellpadding="-2" cellspacing="-2"
 background="../Images/add_manager.gif">
  <tr>
    <td valign="top"><table width="400" height="282" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="11" height="77">&nbsp;</td>
        <td width="373">&nbsp;</td>
        <td width="14"><span class="style2"></span></td>
      </tr>
      <tr>
        <td height="144">&nbsp;</td>
        <td valign="top">
  
          <div align="center">
            <form ACTION="modify_manager_PWD.asp" METHOD="POST" name="form1">
            
            <table width="80%" border="1" cellpadding="-2"  cellspacing="-2"
			 bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr>
                <td width="28%" height="27" bgcolor="#809EA4"><div align="center" class="style1">����Ա���ƣ�</div></td>
                <td width="72%">                  <div align="left">
                    &nbsp;
                    <input name="Name" type="text" class="Sytle_auto" id="Name"
					 value="<%=(rs.Fields.Item("Name").Value)%>" readonly="yes">
                    <input name="H_PWD" type="hidden" id="H_PWD"
					 value="<%=(rs.Fields.Item("PWD").Value)%>">
                    <input name="manager" type="hidden" id="manager"
					 value="<%=session("purview")%>">
</div></td>
                </tr>
              <tr>
                <td height="27" bgcolor="#809EA4"><div align="center">
				<span class="style1">&nbsp;����ԭ���룺</span></div></td>
                <td><div align="left">&nbsp;
                    <input name="PWD_S" type="password" class="Sytle_auto" id="PWD_S">
</div></td>
                </tr>
              <tr>
                <td height="27" bgcolor="#809EA4" class="style1">
				<div align="center">���������룺</div></td>
                <td><div align="left">&nbsp;
                    <input name="PWD_New1" type="password" class="Sytle_auto" id="PWD_New1">
</div></td>
              </tr>
              <tr>
                <td height="27" bgcolor="#809EA4" class="style1">
				<div align="center">ȷ�������룺</div></td>
                <td><div align="left">&nbsp;
                    <input name="PWD_New2" type="password" class="Sytle_auto" id="PWD_New2">
</div></td>
              </tr>
			  <tr>
                <td height="40" colspan="2"><div align="center">
                  <input name="Button" type="button" class="Style_button" value="����"
				   onClick="mycheck()">
                  <input name="Button" type="button" class="Style_button" value="�ر�"
				   onClick="javascrip:opener.location.reload();window.close()">
                </div></td>
                </tr>
              </table>
            </form>
            </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
