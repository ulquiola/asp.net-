<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"--><!--�������ݿ������ļ�-->
<%
TypeID=request.QueryString("TypeID")								'��ȡTypeID��ֵ
if request.Form("Submit")<>"" Then									'�жϱ��Ƿ��ύ
	TypeID=request.Form("TypeID")									'��ȡ��Ԫ��TypeID��ֵ������TypeID����
	owner=request.Form("owner")										'��ȡ��Ԫ��owner��ֵ������owner����
	if TypeID<>"" Then												'�ж�TypeID�Ƿ�Ϊ��
		sql="Update tb_Type set owner="&owner&" where ID="&TypeID	'����ָ����¼
		conn.execute(sql)											'ִ��sql���
		response.Write("<script language='javascript'>alert('��Ϣ�޸ĳɹ���');opener.location.reload();window.close();</script>")
		'������ʾ��Ϣ�Ի���
		response.End()
		'�����������Խű������в���������ظ������
	End If
Else
	set rs=Server.CreateObject("ADODB.RecordSet")					'������¼��
	sql="Select * From tb_user where Status='����'"					'��ѯ����
	rs.open sql,conn,1,3											'�򿪼�¼��
	set rs_T=Server.CreateObject("ADODB.RecordSet")					'������¼��
	sql="Select * From tb_Type where ID="&TypeID					'��ѯ����
	rs_T.open sql,conn,1,3											'�򿪼�¼��
End If
%>
<html>
<head>
<title>��������Ϣ�޸�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="./Css/style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.STYLE1 {
	color: #000000;
	font-size: 9pt;
}
-->
</style></head>
<body>
<table width="240" height="139"  border="0" align="center" cellpadding="-2" cellspacing="-2">
  <tr>
    <td height="139" align="center"><form name="form_U" method="post" action="modify_Type.asp">
        <table width="259" height="124" bgcolor="#FFFFFF" border="0" align="center" cellpadding="-2" cellspacing="-2" class="tabBorder">
          <tr align="center" valign="middle">
            <td height="25" colspan="2" background="../Images/bg_Login.GIF"><span class="STYLE1">=== ��������Ϣ === </span> </td>
          </tr>
          <tr>
            <td width="78" height="40" align="right" valign="middle" class="STYLE1">���������ƣ�</td>
            <td width="168" valign="middle"><input name="TypeName" type="text" id="TypeName" value="<%=rs_T("TypeName")%>" maxlength="20" readonly="yes"></td>
          </tr>
          <tr>
            <td height="30" align="right" valign="middle" class="STYLE1">�������ƣ�</td>
            <td width="168" valign="middle"><span class="word_grey">
              <select name="owner">
                <%For i=1 to rs.recordcount%>
                <option value="<%=rs("ID")%>" <%if rs("ID")=rs_T("owner") then %>selected <%end if%>><%=rs("UserName")%></option>
                <%
			  rs.movenext
			  Next
			  %>
              </select>
              <input name="TypeID" type="hidden" id="TypeID" value="<%=TypeID%>">
</span></td>
          </tr>
          <tr align="center" valign="middle">
            <td height="27" colspan="2"><input name="Submit" type="submit" class="btn_grey" value="����">
&nbsp;
            <input name="Submit2" type="button" class="btn_grey" value="�ر�" onClick="window.close()"></td>
          </tr>
        </table>
    </form></td>
  </tr>
</table>
</body>
</html>
