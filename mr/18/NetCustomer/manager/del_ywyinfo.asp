<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<!--#include file="../../NetCustomer/F_Display.asp" -->
<%''������ʾҵ��Ա��Ϣ�ļ�¼��
If (Request.QueryString("ywynumber") <> "") Then 
  session("NO") = Request.QueryString("ywynumber")
End If
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM DB_WorkerInfo WHERE ywynumber = '"&session("NO")& "'"
rs.open sql,conn,1,3%>
<%'ɾ��ҵ��Ա��Ϣ
if request.Form("name")<>"" then
	Del="Delete from DB_WorkerInfo where ywynumber = '"&session("NO")& "'"
	conn.execute(Del)%>
	<script language="javascript">
	alert("����ɾ���ɹ���");
	opener.location.href="manage_ywy.asp";//����������ҳ
	window.close();
	</script>
<% end if%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>ҵ��Ա��Ϣ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image: url(../Images/bg.gif);
}
.style1 {color: #FFFFFF}
.style2 {color: #a2bcc5}
-->
</style></head>

<body>
<table width="595" height="242" border="0" cellpadding="-2" cellspacing="-2" background="../Images/manage.gif">
  <tr>
    <td valign="top"><table width="595" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="10" height="112">&nbsp;</td>
        <td width="577">&nbsp;</td>
        <td width="10"><span class="style2"></span></td>
      </tr>
      <tr>
        <td height="258">&nbsp;</td>
        <td valign="top"><div align="center">
          <form ACTION="del_ywyinfo.asp" METHOD="POST" name="form1">
            <table border="1" align="center" cellpadding="-2" cellspacing="-2" bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr valign="middle">
                <td width="72" align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">�ࡡ���ţ�</div></td>
                <td width="105">                    <div align="left">
                  <input name="ywynumber" type="text" class="Sytle_text" id="ywynumber" value="<%=(rs.Fields.Item("ywynumber").Value)%>" size="32" readonly="yes">
                </div></td>
                <td width="71" bgcolor="#809EA4"><div align="center" class="style1">�ա�������</div></td>
                <td width="96"><div align="left">
                  <input name="name" type="text" class="Sytle_text" value="<%=(rs.Fields.Item("name").Value)%>" size="32">
                </div></td>
                <td width="68" bgcolor="#809EA4"><div align="center" class="style1">�ԡ�����</div></td>
                <td width="97"><div align="left">
                  <select name="sex" id="select3">
                    <option value="��" <%If (Not isNull((rs.Fields.Item("sex").Value))) Then If ("��" = CStr((rs.Fields.Item("sex").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>��</option>
                    <option value="Ů" <%If (Not isNull((rs.Fields.Item("sex").Value))) Then If ("Ů" = CStr((rs.Fields.Item("sex").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>Ů</option>
                  </select>
                </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">�񡡡��壺</div></td>
                <td><div align="left">
                  <input name="folk" type="text" class="Sytle_text" value="<%=(rs.Fields.Item("folk").Value)%>" size="32">
                </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">�������գ�</div></td>
                <td><div align="left">
                  <input name="birthday" type="text" class="Sytle_text" value="<%=(rs.Fields.Item("birthday").Value)%>" size="32">
                </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">�볧���ڣ�</div></td>
                <td><div align="left">
                  <input name="rcdate" type="text" class="Sytle_text" value="<%=(rs.Fields.Item("rcdate").Value)%>" size="32">
                </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">ѧ��������</div></td>
                <td><div align="left">
                  <select name="xl" id="select4">
                    <option value="��ʿ������" selected <%If (Not isNull((rs.Fields.Item("xl").Value))) Then If ("��ʿ������" = CStr((rs.Fields.Item("xl").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>��ʿ������</option>
                    <option value="˶ʿ" <%If (Not isNull((rs.Fields.Item("xl").Value))) Then If ("˶ʿ" = CStr((rs.Fields.Item("xl").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>˶ʿ</option>
                    <option value="�о���" <%If (Not isNull((rs.Fields.Item("xl").Value))) Then If ("�о���" = CStr((rs.Fields.Item("xl").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>�о���</option>
                    <option value="����" <%If (Not isNull((rs.Fields.Item("xl").Value))) Then If ("����" = CStr((rs.Fields.Item("xl").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>����</option>
                    <option value="ר��" <%If (Not isNull((rs.Fields.Item("xl").Value))) Then If ("ר��" = CStr((rs.Fields.Item("xl").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>ר��</option>
                    <option value="���м���ר" <%If (Not isNull((rs.Fields.Item("xl").Value))) Then If ("���м���ר" = CStr((rs.Fields.Item("xl").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>���м���ר</option>
                    <option value="���м�����" <%If (Not isNull((rs.Fields.Item("xl").Value))) Then If ("���м�����" = CStr((rs.Fields.Item("xl").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>���м�����</option>
                  </select>
</div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">���֤�ţ�</div></td>
                <td colspan="3"><div align="left">
                  <input name="sfzNO" type="text" class="Sytle_auto" value="<%=(rs.Fields.Item("sfzNO").Value)%>" size="32">
                </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">ר����ҵ��</div></td>
                <td><div align="left">
                  <select name="zy" id="select">
                    <option value="�����" selected <%If (Not isNull((rs.Fields.Item("zy").Value))) Then If ("�����" = CStr((rs.Fields.Item("zy").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>�����</option>
                    <option value="ʳƷ" <%If (Not isNull((rs.Fields.Item("zy").Value))) Then If ("ʳƷ" = CStr((rs.Fields.Item("zy").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>ʳƷ</option>
                    <option value="��е����" <%If (Not isNull((rs.Fields.Item("zy").Value))) Then If ("��е����" = CStr((rs.Fields.Item("zy").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>��е����</option>
                    <option value="ҽ��" <%If (Not isNull((rs.Fields.Item("zy").Value))) Then If ("ҽ��" = CStr((rs.Fields.Item("zy").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>ҽ��</option>
                    <option value="��ʿ" <%If (Not isNull((rs.Fields.Item("zy").Value))) Then If ("��ʿ" = CStr((rs.Fields.Item("zy").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>��ʿ</option>
                    <option value="��ʦ" <%If (Not isNull((rs.Fields.Item("zy").Value))) Then If ("��ʦ" = CStr((rs.Fields.Item("zy").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>��ʦ</option>
                    <option value="����" <%If (Not isNull((rs.Fields.Item("zy").Value))) Then If ("����" = CStr((rs.Fields.Item("zy").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>����</option>
                  </select>
                </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">��ϵ�绰��</div></td>
                <td colspan="3"><div align="left">
                  <input name="tel" type="text" class="Sytle_auto" value="<%=(rs.Fields.Item("tel").Value)%>" size="32">
                </div></td>
                </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">��������</div></td>
                <td colspan="5"><div align="left">                  <table width="98%"  cellspacing="-2" cellpadding="-2">
                    <tr>
                      <td width="86%">
						<% qyNO=rs("workqy")
						call Display(qyNO)  '��ʾ��������%>
					  </td>
                      <td width="14%"><input name="workqy" type="hidden" id="workqy4" value="<%=rs("workqy")%>"></td>
                    </tr>
                  </table>
</div>                  <div align="left"></div></td>
                </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">�ء���ַ��</div></td>
                <td colspan="5">
                  <div align="left">
                    <input name="address" type="text" class="Style_upload" value="<%=(rs.Fields.Item("address").Value)%>" size="32">                  
                  </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">Email:</div></td>
                <td colspan="5">
                  <div align="left">
                    <input name="email" type="text" class="Sytle_auto" value="<%=(rs.Fields.Item("email").Value)%>" size="32">                
                  </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">������ע��</div></td>
                <td colspan="5">
                  <div align="left">
                    <input name="memo" type="text" class="Style_upload" value="<%=(rs.Fields.Item("memo").Value)%>" size="32">                
                  </div></td>
              </tr>
              <tr valign="middle">
                <td height="35" colspan="6" nowrap><div align="center">
                    <input name="Submit" type="submit" class="Style_button" value="ɾ��">                
                    <input name="Submit2" type="button" class="Style_button" value="�ر�" onClick="JScript:window.close()">
                </div></td>
                </tr>
            </table>
          </form>
          <p>&nbsp;</p>
        </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="21">&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
