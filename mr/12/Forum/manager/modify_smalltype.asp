<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"-->
<%
ID=request.QueryString("ID")					'��ȡ��Ԫ��ID��ֵ������ID����
if request.Form("Submit")<>"" Then				'�жϱ��Ƿ��ύ
	ID=request.Form("ID")						'��ȡ��Ԫ��ID��ֵ������ID����
	typeid=request.Form("typeid")				'��ȡ��Ԫ��typeid��ֵ������typeid����
	smalltypename=request.Form("smalltypename")	'��ȡ��Ԫ��smalltypename��ֵ������smalltypename����
	userid=request.Form("userid")				'��ȡ��Ԫ��userid��ֵ������userid����
	account=request.Form("account")				'��ȡ��Ԫ��account��ֵ������account����
	if ID<>"" Then								'�ж�IDֵ�Ƿ�Ϊ��
		sql="Update tb_smalltype set typeid="&typeid&",smalltypename='"&smalltypename&"',userid="&userid&",account='"&account&"'where ID="&ID
		'Ӧ��Update������ָ����������Ϣ
		conn.execute(sql)		'ִ��sql���
		response.Write("<script language='javascript'>alert('��Ϣ�޸ĳɹ���');opener.location.reload();window.close();</script>")
		'������ʾ�Ի���
		response.End()
		'�����������Խű������в���������ظ������
	End If
Else
	set rs=Server.CreateObject("ADODB.RecordSet")		'������¼��
	sql="Select * From tb_user where Status='����'"		'��ѯ������Ϣ
	rs.open sql,conn,1,3								'�򿪼�¼��
	set rs_T=Server.CreateObject("ADODB.RecordSet")		'������¼��
	sql="Select * From tb_type"							'��ѯ����
	rs_T.open sql,conn,1,3								'�򿪼�¼��
	set rs_1=Server.CreateObject("ADODB.RecordSet")		'������¼��
	sql1="Select * From view_smalltype where ID="&ID	'��ѯ����
	rs_1.open sql1,conn,1,3								'�򿪼�¼��
End If
%>		
<html>
<head>
<title>�����Ϣ�޸�</title>
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
.STYLE1 {
	color: #000000;
	font-size: 9pt;
}
.STYLE2 {
	font-size: 10pt;
	color: #000000;
}
-->
</style></head>
<body>
<table width="496" height="205"  border="0" cellpadding="-2" cellspacing="-2">
  <tr>
    <td height="139" align="center" valign="top"><form name="form_U" method="post" action="modify_smalltype.asp">
        <table width="507" height="198" bgcolor="#FFFFFF" border="0" align="center" cellpadding="-2" cellspacing="-2" class="tabBorder">
          <tr align="center" valign="middle">
            <td height="21" background="../Images/bg_Login.GIF"><div align="center" class="STYLE2"></div></td>
            <td height="21" background="../Images/bg_Login.GIF"><div align="left"><span class="STYLE2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;=== �����Ϣ === </span></div></td>
          </tr>
          <tr>
            <td width="75" height="37" align="right" valign="middle" class="STYLE1">������������</td>
            <td width="432" valign="middle">
			<select name="typeid" id="typeid">
			<%For i=1 to rs_T.recordcount%>
            <option value="<%=rs_T("ID")%>" <%if rs_T("ID")=rs_1("typeid") then %>selected <%end if%>><%=rs_T("TypeName")%></option>
             <%
			  rs_T.movenext
			 Next
			  %>
            </select>
			</td>
          </tr>
          <tr>
            <td width="75" height="41" align="right" valign="middle" class="STYLE1">������ƣ�</td>
            <td valign="middle"><input name="smalltypename" type="text" class="inputcss1" id="smalltypename" size="30" maxlength="20" value="<%=rs_1("smalltypename")%>"></td>
          </tr>
          <tr>
            <td width="75" height="30" align="right" valign="middle" class="STYLE1">������</td>
            <td valign="middle"><span class="word_grey">
              <select name="userid" id="userid">
			<%For i=1 to rs.recordcount%>
            <option value="<%=rs("ID")%>" <%if rs("ID")=rs_1("userid") then %>selected <%end if%>><%=rs("username")%></option>
             <%
			  rs.movenext
			 Next
			  %>
            </select>              
            </span>
			</td>
          </tr>
          <tr>
            <td width="75" height="5" align="right" valign="middle" class="STYLE1">���˵����            </td>
            <td valign="middle"><span class="STYLE1">
             <textarea name="account" cols="40" rows="3" id="account" value="<%=rs_1("account")%>"><%=rs_1("account")%></textarea>
              <span class="word_grey">
              <input name="ID" type="hidden" id="ID" value="<%=ID%>">
            </span></span></td>
          </tr>
          <tr align="center" valign="middle">
            <td height="78" colspan="2">
                        <div align="left">
                          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                          <input name="Submit" type="submit" class="btn_grey" value="����">
                          &nbsp;
                          <input name="Submit2" type="button" class="btn_grey" value="�ر�" onClick="window.close()">
                        </div></td></tr>
        </table>
    </form></td>
  </tr>
</table>
</body>
</html>
