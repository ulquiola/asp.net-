<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp" -->
<%
if request.Form("UserName")<>"" then
	DEL="Delete From Tab_User where ID='"&session("ID")&"'"
	conn.execute(Del)%>
	<script language="javascript">
	alert("Ա����Ϣ�Ѿ�ɾ����");
	opener.location.reload();
	window.close();
	</script>
<%end if%>
<%
if Request.QueryString("ID")<>""then
	session("ID")=Request.QueryString("ID")
end if
Set rs_personnel = Server.CreateObject("ADODB.Recordset")
sql_P="SELECT ID,UserName,Name,PWD,purview,branch,sex,Email,Tel,Address"&_
" FROM Tab_User WHERE ID="&session("ID")&""
rs_personnel.open sql_P,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>ɾ��Ա����Ϣ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
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
    <td width="817" height="403" valign="top"><table width="100%" height="89"  border="0" cellpadding="-2" cellspacing="-2">
      <tr>
        <td height="41" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE4">ɾ��Ա����Ϣ</span></td>
      </tr>
      <tr>
        <td height="44">&nbsp;</td>
      </tr>
    </table>      
      <br>
      <form ACTION="personnel_del.asp" METHOD="POST" name="form1">
        <table width="95%" height="222"  border="0" align="center" cellpadding="-2" cellspacing="-2">
          <tr>
            <td colspan="6"><span class="style1">&nbsp;&nbsp;&nbsp;�û�����</span>&nbsp;<%=(rs_personnel("UserName"))%></td>
          </tr>
		  <tr>
            <td width="16%" height="27"><div align="center" class="style1">������</div></td>
            <td width="39%" height="27"><span class="style1">
              <input name="username" type="text" class="Sytle_text" id="username2" readonly value="<%=(rs_personnel("Name"))%>">
            </span> </td>
            <td width="9%" height="27" class="style1">Ȩ�ޣ�</td>
            <td width="13%" height="27"><select name="purview" id="purview">
              <option value="ϵͳ" <%If (Not isNull((rs_personnel("purview")))) Then If ("ϵͳ" = CStr((rs_personnel("purview")))) Then Response.Write("SELECTED") : Response.Write("")%>>ϵͳ</option>
              <option value="��д" <%If (Not isNull((rs_personnel("purview")))) Then If ("��д" = CStr((rs_personnel("purview")))) Then Response.Write("SELECTED") : Response.Write("")%>>��д</option>
              <option value="ֻ��" <%If (Not isNull((rs_personnel("purview")))) Then If ("ֻ��" = CStr((rs_personnel("purview")))) Then Response.Write("SELECTED") : Response.Write("")%>>ֻ��</option>
            </select></td>
            <td width="10%" height="27"><span class="style1">�Ա�</span></td>
            <td width="13%" height="27"><span class="style1">
              <select name="sex" id="select2">
                <option value="��" <%If (Not isNull((rs_personnel("sex")))) Then If ("��" = CStr((rs_personnel("sex")))) Then Response.Write("SELECTED") : Response.Write("")%>>��</option>
                <option value="Ů" <%If (Not isNull((rs_personnel("sex")))) Then If ("Ů" = CStr((rs_personnel("sex")))) Then Response.Write("SELECTED") : Response.Write("")%>>Ů</option>
              </select>
            </span></td>
          </tr>
          <tr>
            <td height="27"><div align="center" class="style1">�绰��</div></td>
            <td height="27"><input name="tel" type="text" class="Sytle_text" id="tel" value="<%=(rs_personnel("Tel"))%>"></td>
            <td height="27" class="style1">���ţ�</td>
            <td height="27" colspan="3"><select name="branch" id="branch">
              <option value="������" <%If (Not isNull((rs_personnel("branch")))) Then If ("������" = CStr((rs_personnel("branch")))) Then Response.Write("SELECTED") : Response.Write("")%>>������</option>
              <option value="���²�" <%If (Not isNull((rs_personnel("branch")))) Then If ("���²�" = CStr((rs_personnel("branch")))) Then Response.Write("SELECTED") : Response.Write("")%>>���²�</option>
              <option value="���۲�" <%If (Not isNull((rs_personnel("branch")))) Then If ("���۲�" = CStr((rs_personnel("branch")))) Then Response.Write("SELECTED") : Response.Write("")%>>���۲�</option>
            </select></td>
          </tr>
          <tr>
            <td height="27" class="style1"><div align="center">Email��</div></td>
            <td height="27" colspan="5"><input name="email" type="text" class="Style_subject" id="email" value="<%=(rs_personnel("Email"))%>"></td>
          </tr>
          <tr>
            <td height="27" class="style1"><div align="center">��ַ��</div></td>
            <td height="27" colspan="5"><input name="address" type="text" class="Style_subject" id="address" value="<%=(rs_personnel("Address"))%>"></td>
          </tr>
          <tr>
            <td colspan="6"><div align="center">
                <br>
                <input name="Button" type="submit" class="Style_button_del" value="ȷ��ɾ��">
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