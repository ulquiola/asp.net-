<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/Validate.asp"-->
<!--#include file="Conn/conn.asp"-->
<%
function level(num)
	select case num
		case "1"
			level = "һ������"
		case "2"
			level = "��������"
		case "3"
			level = "��������"
		case "4"
			level = "�ļ�����"
		case "5"
			level = "�弶����"
		case "6"
			level = "��������"
		case "7"
			level = "�߼�����"
		case "8"
			level = "�˼�����"
	end select
end function
getkey=Session("StuID")
set rs = server.createobject("adodb.recordset") 
rssql = "Select * from tb_StuResult where ����='"&getkey&"'"
rs.open rssql,conn,1,3
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ѯ�ɼ�</title>
<link rel="stylesheet" type="text/css" href="Css/style.css">
<style type="text/css">
<!--
.STYLE1 {
	color: #333333;
	font-weight: bold;
}
body {
	margin-top: 0px;
}
-->
</style>
</head>
<body leftmargin="0">
<table cellpadding="0" cellspacing="0" border="0" width="100%" height="90%">
  <tr>
    <td align="center" height="18"><span class="STYLE1">��ѯ�ɼ�</span> </td>
  </tr>
  <tr>
    <td height="354" valign="top"><table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
        <tr>
          <td valign="top">
            <table width="100%"  border="1" cellspacing="0" cellpadding="0" align="center"
 bordercolor="#FFFFFF" bordercolordark="#1985DB" bordercolorlight="#FFFFFF">
              <tr>
                <td height="26" colspan="6"  align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="26%" height="24" align="center" bgcolor="#00CCFF"><span style="color:black;"><strong>����ID��</strong></span></td>
                    <td width="40%" align="center">
						<strong><span style="color:black;">
						<%
							if not rs.eof then
						 		response.Write(rs("����ID��"))
							else
								response.Write("<font color=red>û�з��������ļ�¼��</font>")
								response.End()
							end if
						%></span></strong></td>
                    <td width="34%">&nbsp;</td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td width="26%" height="24"  align="center" nowrap bgcolor="#00CCFF" style="color:black;"><strong>����</strong></td>
                <td width="11%" height="26"  align="center" nowrap bgcolor="#00CCFF" style="color:black;">����</td>
                <td width="19%" height="26"  align="center" nowrap bgcolor="#00CCFF" style="color:black;">����</td>
                <td width="16%" height="26"  align="center" nowrap bgcolor="#00CCFF" style="color:black;">�Ա�</td>
                <td width="14%" height="26"  align="center" nowrap bgcolor="#00CCFF" style="color:black;">�꼶</td>
                <td width="14%" height="26"  align="center" nowrap bgcolor="#00CCFF" style="color:black;">�������</td>
              </tr>
              <tr>
                <td height="24"  align=center nowrap><%=rs("����")%></td>
                <td height="24"  align=center nowrap style="color:black;"><%=rs("����")%>&nbsp;</td>
                <td height="24"  align=center nowrap style="color:black;"><%=rs("����")%>&nbsp;</td>
                <td height="24"  align=center nowrap style="color:black;"><%=rs("�Ա�")%>&nbsp;</td>
                <td height="24"  align=center nowrap style="color:black;"><%=rs("�꼶")%>&nbsp;</td>
                <td height="24"  align=center nowrap style="color:black;"><%=rs("�������")%>&nbsp;</td>
              </tr>
            </table>
            <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#1985DB">
              <tr>
                <td width="7%" height="24" align="center" bgcolor="#00CCFF">��Ŀ</td>
                <td width="19%" align="center" bgcolor="#00CCFF">�÷�</td>
                <td width="11%" align="center" bgcolor="#00CCFF">�ȼ�</td>
                <td width="63%" align="center" bgcolor="#00CCFF">����</td>
              </tr>
             
                <%
				'�����
				set rs1 = server.createobject("adodb.recordset") 
				rssql = "SELECT tb_StuResult.*, tb_ear."&level(rs("����"))&" as ��������, tb_talk."&level(rs("����"))&" as ��������, tb_write."&level(rs("����"))&" as ��������, tb_zonghe."&level(rs("����"))&" as �ۺ����� "&_
				"FROM tb_zonghe INNER JOIN ("&_
					"tb_write INNER JOIN ("&_
						"tb_talk INNER JOIN ("&_
							"tb_ear INNER JOIN tb_StuResult ON tb_ear.�����ȼ� = tb_StuResult.�����ȼ�"&_
						") ON tb_talk.����ȼ� = tb_StuResult.����ȼ�"&_
					") ON tb_write.���ĵȼ� = tb_StuResult.���ĵȼ�"&_
				") ON tb_zonghe.�ۺϵȼ� = tb_StuResult.�ۺϵȼ�"&_
				" where tb_sturesult.���� = '130008100002'"
				rs1.open rssql,conn,1,3
				%>
			<tr>
                <td align="center" bgcolor="#00CCFF">����</td>
                <td align="center"><%=rs1("����÷�")%>&nbsp;</td>
                <td align="center"><%=server.htmlencode(rs1("����ȼ�"))%>&nbsp;</td>
                <td><%=rs1("��������")%>&nbsp;</td>
              </tr>
              <tr>
                <td align="center" bgcolor="#00CCFF">����</td>
                <td align="center"><%=server.htmlencode(rs1("�����÷�"))%>&nbsp;</td>
                <td align="center"><%=server.htmlencode(rs1("�����ȼ�"))%>&nbsp;</td>
                <td><%=rs1("��������")%></td>
              </tr>
                <td align="center" bgcolor="#00CCFF">�ۺ�</td>
                <td align="center" bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="1" cellpadding="0">
                    <tr>
                      <td width="49%" align="center" bgcolor="#FFFFFF"><span style="color:black;">�ۺ����</span></td>
                      <td width="22%" align="center" bgcolor="#FFFFFF"><span style="color:black;"><%=rs1("�ۺ����")%></span>&nbsp;</td>
                      <td width="29%" rowspan="5" align="center" bgcolor="#FFFFFF"><%=server.htmlencode(rs1("�ۺϵ÷�"))%>&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="center" bgcolor="#FFFFFF"><span style="color:black;">�ۺ����</span></td>
                      <td align="center" bgcolor="#FFFFFF"><span style="color:black;"><%=rs1("�ۺ����")%></span>&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="center" bgcolor="#FFFFFF"><span style="color:black;">�ۺ����</span></td>
                      <td align="center" bgcolor="#FFFFFF"><span style="color:black;"><%=rs1("�ۺ����")%></span>&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="center" bgcolor="#FFFFFF"><span style="color:black;">�ۺ����</span></td>
                      <td align="center" bgcolor="#FFFFFF"><span style="color:black;"><%=rs1("�ۺ����")%></span>&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="center" bgcolor="#FFFFFF"><span style="color:black;">�ۺ����</span></td>
                      <td align="center" bgcolor="#FFFFFF"><span style="color:black;"><%=rs1("�ۺ����")%></span>&nbsp;</td>
                    </tr>
                  </table></td>
                <td align="center"><%=server.htmlencode(rs1("�ۺϵȼ�"))%>&nbsp;</td>
                <td><%=rs1("�ۺ�����")%>
                &nbsp;</td>
              </tr>
              <tr>
                <td align="center" bgcolor="#00CCFF">����</td>
                <td align="center"><%=rs1("���ĵ÷�")%>&nbsp;</td>
                <td align="center"><%=server.htmlencode(rs1("���ĵȼ�"))%>&nbsp;</td>
                <td><%=rs1("��������")%>
                &nbsp;</td>
              </tr>
              <tr>
                <td height="26" align="center" bgcolor="#00CCFF"><strong>�ܳɼ�</strong></td>
                <td align="center"><strong><%=server.htmlencode(rs1("�ܷ�"))%></strong>&nbsp;</td>
                <td align="center"><strong><%=server.htmlencode(rs1("����"))%></strong>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
          </table></td>
        </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
