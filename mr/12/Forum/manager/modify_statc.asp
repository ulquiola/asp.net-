<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"--><!--�������ݿ������ļ�-->
<!--#include file="../head.asp"-->
<%
If Request.QueryString("ID")<>""then			'�жϽ��յ�IDֵ�Ƿ���ڿ�
session("ID")=Request.QueryString("ID")			'Ϊsession("ID")������ֵ
end if
Set rs_personnel = Server.CreateObject("ADODB.Recordset")	'������¼��
sql_P="SELECT UserName,TrueName,sex,birthday,Tel,qq,ICO,Email,homepage,address,Status FROM tb_User where ID="&session("ID")&""
'��ѯ����
rs_personnel.open sql_p,conn,1,3				'�򿪼�¼��
%>
<%
if request.Form("UserName")<>"" then			'�жϽ��յ�UserNameֵ�Ƿ���ڿ�
	Status=request.Form("Status")				'ΪStatus������ֵ
	UP="Update tb_user set Status='"&Status&"' where ID="&session("ID")&""	'����ָ���ļ�¼
	conn.execute(UP)														'ִ��UP���
	%>
<script langeuae="javascript">
alert("�û���Ϣ�޸ĳɹ�����")					//������ʾ�Ի���
window.location.href="User_Manage.asp";			//��ת��ָ��ASPҳ��
</script>
<%end if%>
<%
If Request.QueryString("id") <> "" then		'�жϽ��յ�IDֵ�Ƿ���ڿ�
session("id")=Request.QueryString("id")		'Ϊsession("ID")������ֵ
end if
Set rs1 = Server.CreateObject("ADODB.Recordset")	'������¼��
sql1="SELECT UserName,TrueName,sex,birthday,Tel,qq,ICO,Email,homepage,address,Status FROM tb_User where ID="&session("ID")&""
'��ѯ����
rs1.open sql1,conn,1,3								'�򿪼�¼��
%>
<html>
<head>
<title>�޸��û���Ϣ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Css/style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.STYLE1 {	font-size: 10pt;
	font-weight: bold;
}
-->
</style></head>
<body bgcolor="B9DFFF">
<table width="777" height="341"  border="0" align="center" cellpadding="-2" cellspacing="-2">
  <tr>
    <td valign="top"><table width="100%" border="0" cellpadding="-2" cellspacing="-2" class="tableBorder">
      <tr>
        <td></td>
      </tr>
      <tr>
        <td></td>
      </tr>
      <tr>
        <td height="2"></td>
      </tr>
      <tr>
        <td valign="top">
      <table width="100%" height="205"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder">
        <tr>
          <td width="13" height="32" bgcolor="#4EAEE1">&nbsp;</td>
          <td width="689" bgcolor="#4EAEE1" style="color:#FFFFFF">��<span class="STYLE1"><font color="#FFFFFF">��ǰλ�ã��鿴�û���Ϣ</font></span> �� </td>
          <td width="73" bgcolor="#4EAEE1">&nbsp;</td>
        </tr>
        <tr valign="top">
          <td height="171" colspan="3"><table width="100%" height="129"  border="0" cellpadding="-2" cellspacing="-2">
            <tr>
              <td valign="top">
	  			  <table width="100%" height="265"  border="0" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF" class="tableBorder_Top">
			  <tr>
			  <td height="5"></td>
			  </tr>
                <tr>
                 <td width="700" valign="top"><table width="90%"  border="0" align="center" cellpadding="-2" cellspacing="-2">
                    <tr>
                      <td><form action="modify_statc.asp" method="post" name="myform">
                        <table width="100%"  border="0" cellspacing="-2" cellpadding="-2">
                          <tr>
                            <td width="18%" height="30" align="center">�� �� ����</td>
                            <td width="82%"><input name="UserName" type="text" class="inputcss1" id="UserName4" value="<%=rs_personnel("UserName")%>" maxlength="20" readonly="yes">
      *      </td>
                            </tr>
                          <tr>
                            <td height="28" align="center">��ʵ������</td>
                            <td height="28"><input name="TrueName" type="text" class="inputcss1" id="TrueName4" value="<%=rs_personnel("TrueName")%>" maxlength="10" readonly="yes">
                              *</td>
                            </tr>
                          <tr>
                            <td height="28" align="center">��&nbsp;&nbsp;&nbsp;&nbsp;��</td>
                            <td><input name="sex" type="radio" class="noborder" value="��" <%If rs_personnel("Sex")="��" Then%>checked<%End If%> readonly="yes">
                              ��&nbsp;
  <input name="sex" type="radio" class="noborder" value="Ů"<%If rs_personnel("Sex")="Ů" Then %>checked<%End If%> readonly="yes">
  Ů</td></tr>
                          <tr>
                            <td height="28" align="center">��&nbsp;&nbsp;&nbsp;&nbsp;�գ�</td>
                            <td class="word_grey"><input name="birthday" type="text" class="inputcss1" id="Tel" value="<%=rs_personnel("birthday")%>" readonly="yes">
                              *�����ڸ�ʽΪ��yyyy-mm-dd �磺1980-07-17��</td>
                          </tr>
                          <tr>
                            <td height="28" align="center">��ϵ�绰��</td>
                            <td><input name="Tel" type="text" class="inputcss1" id="Tel2" value="<%=rs_personnel("Tel")%>" readonly="yes"></td>
                            </tr>
                          <tr>
                            <td height="28" align="center">OICQ���룺</td>
                            <td><input name="qq" type="text" class="inputcss1" id="qq2" value="<%=rs_personnel("QQ")%>" readonly="yes"></td>
                            </tr>
                          <tr>
                            <td height="62" align="center">ѡ��ͷ��</td>
                            <td class="word_grey"><table width="316" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="13" height="47"></td>
                                <td width="145"><img src="../Images/head/<%=rs_personnel("ICO")%>" name="img" width="60" height="60" border="1"></td>
                                <td width="225">&nbsp;</td>
                              </tr>
                            </table></td>
                          </tr>
                          <tr>
                            <td height="28" align="center" style="padding-left:10px">Email��</td>
                            <td class="word_grey"><input name="Email" type="text" class="inputcss1" id="PWD224" value="<%=rs_personnel("Email")%>" size="50" readonly="yes">
                              *</td>
                          </tr>
                          <tr>
                            <td height="28" align="center">������ҳ��</td>
                            <td class="word_grey"><input name="homepage" type="text" class="inputcss1" id="homepage" value="<%=rs_personnel("homepage")%>" size="50" readonly="yes"></td>
                          </tr>
                          <tr>
                            <td height="23" align="center">��ͥסַ��</td>
                            <td class="word_grey"><input name="address" type="text" class="inputcss1" id="address" value="<%=rs_personnel("Address")%>" size="50" readonly="yes"			></td>
                          </tr>
                          <tr>
                            <td height="32" align="center">�û�Ȩ�ޣ�</td>
                            <td class="word_grey">							
							<select name="Status" class="text" id="Status">
      <option value="��ͨ�û�" <%if(InStr(rs1("Status"),"��ͨ�û�") > 0)Then%>selected<%End IF%>>��ͨ�û�</option>
      <option value="����" <%If(InStr(rs1("Status"),"����") > 0)Then%>selected<%End IF%>>����</option>
    </select>							</td>
                          </tr>
                          <tr>
                            <td height="34">&nbsp;</td>
                            <td class="word_grey"><input name="submit" type="submit" class="btn_grey" value="ȷ������">
                                <input name="Submit2" type="reset" class="btn_grey" value="������д">
                                
                                <input name="Submit22" type="button" class="btn_grey" value="���ع���ҳ" onClick="Jscript:window.location.href='User_Manage.asp';"></td>
                          </tr>
                        </table>
					  
					  </form></td>
                    </tr>
                  </table></td>
                </tr>
				<tr>
				<td height="5"></td>
				</tr>
              </table>			  </td>
            </tr>
          </table>
		  </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
