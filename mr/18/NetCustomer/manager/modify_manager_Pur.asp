<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<%
if request("ID")<>"" then
session("ID")=request("ID")
end if
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM DB_manager WHERE ID = " + Replace(session("ID"), "'", "''") + ""
rs.open sql,conn,1,3
if not rs.eof then
	if request.Form("purview")<>"" then
	purview=request.Form("purview")
	Up="Update DB_manager set purview='"& purview&"' WHERE ID = " +_
	 Replace(session("ID"), "'", "''") + ""
	conn.execute(UP)%>
	<script language="javascript">
	alert("Ȩ���޸ĳɹ���");
	window.close();
	</script>
	<%end if
end if
%>
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
if(form1.PWD.value=="")
{alert("���������Ա���룡");form1.PWD.focus();return;}
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
            <form ACTION="modify_manager_Pur.asp" METHOD="POST" name="form1">
            
            <table width="80%" border="1" cellpadding="-2"  cellspacing="-2"
			 bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr>
                <td width="28%" height="27" bgcolor="#809EA4"><div align="center"
				 class="style1">����Ա���ƣ�</div></td>
                <td width="72%">
                  <div align="left">
                    &nbsp;
                    <input name="Name" type="text" class="Sytle_auto" id="Name"
					 value="<%=(rs.Fields.Item("Name").Value)%>" readonly="yes">
                  </div></td>
                </tr>
              <tr>
                <td height="27" bgcolor="#809EA4"><div align="center"><span class="style1">
				&nbsp;����Ա���룺</span></div></td>
                <td><div align="left">&nbsp;
                    <input name="PWD" type="password" class="Sytle_auto" id="PWD"
					 readonly="yes" value="<%=(rs.Fields.Item("PWD").Value)%>">
</div></td>
                </tr>
              <tr>
                <td height="27" bgcolor="#809EA4" class="style1"><div align="center">
				����ԱȨ�ޣ�</div></td>
                <td><div align="left">&nbsp;
                  <select name="purview" id="purview">
                    <option value="ϵͳ" <%If (Not isNull(rs("purview"))) Then
					 If ("ϵͳ" = CStr(rs("purview"))) Then
					  Response.Write("SELECTED") : Response.Write("")
					 end if
					 end if%>>ϵͳ������</option>
                    <option value="һ��" <%If (Not isNull(rs("purview"))) Then
					 If ("һ��" = CStr(rs("purview"))) Then
					  Response.Write("SELECTED") : Response.Write("")
					  end if
					  end if%>>һ�㡡����</option>
                  </select>
                </div></td>
              </tr>
              <tr>
                <td height="40" colspan="2"><div align="center">
                  <input name="Button" type="button" class="Style_button" value="����" onClick="mycheck()">
                  <input name="Submit2" type="reset" class="Style_button" value="����">
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