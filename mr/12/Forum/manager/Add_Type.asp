<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"--><!--�������ݿ������ļ�-->
<%
if request.Form("TypeName")<>"" Then	'��ȡTypeNameֵ
	TName=request.Form("TypeName")		'��ȡ��Ԫ��TypeName��ֵ������TName����
	owner=request.Form("selected")		'��ȡ��Ԫ��selected��ֵ������owner����
	if TName<>"" Then					'�ж�TName�Ƿ���ֵ
		set rs=Server.CreateObject("ADODB.RecordSet")				'������¼��
		sql="Select * From tb_Type where TypeName='"&TName&"'"		'��ѯ����
		rs.open sql,conn,1,3										'�򿪼�¼��
		if not(rs.eof and rs.bof) Then								'�ж��Ƿ��м�¼
			response.Write("<script language='javascript'>alert('���������������Ѿ����ڣ����������룡');window.location.href='Add_Type.asp';</script>")
			'������ʾ�Ի���
			response.End()
			'�����������Խű������в���������ظ������
		Else
			sql="Insert Into tb_Type (TypeName,owner) values('"&TName&"',"&owner&")"
			'Ӧ��Insert Into������������Ϣ
			conn.execute(sql)
			'ִ��sql���
			response.Write("<script language='javascript'>alert('��Ϣ��ӳɹ���');window.location.href='Add_Type.asp';</script>")
			'������ʾ�Ի��򣬲���ת��ָ����ASPҳ��
		End If
	Else
		response.Write("<script language='javascript'>alert('�����������������ƣ�')</script>")
		'������ʾ�Ի���
	End If
Else
	set rs=Server.CreateObject("ADODB.RecordSet")			'������¼��
	sql="Select * From tb_user where Status='����'"			'��ѯ����
	rs.open sql,conn,1,3									'�򿪼�¼��
End If
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Css/style.css" rel="stylesheet">
</head>
<script language="javascript">
function mycheck()
{//�����Զ��庯��
if(form_U.TypeName.value=="")
{alert("�����������Ʋ���Ϊ�գ�");form_U.TypeName.focus();return false}
if(form_U.selected.value==0)
{alert("��ѡ���������������İ�����");form_U.selected.focus();return false}
form_U.submit();
}
</script>
<body  bgcolor="B9DFFF">
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#4EAEE1">
  <form name="form_U" method="post" action="Add_Type.asp">
	    <tr>
          <td height="22" colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
            <tr>
              <td height="22" bgcolor="#4EAEE1">&nbsp;&nbsp;<font color="#FFFFFF"><strong>��ǰλ�ã����������</strong></font></td>
            </tr>
          </table></td>
	    </tr>
        <tr>
          <td width="20%" height="22" bgcolor="#FFFFFF"><div align="right">���������ƣ�</div></td>
          <td width="80%" bgcolor="#FFFFFF">&nbsp;
          <input name="TypeName" type="text" class="inputcss1" id="TypeName" maxlength="20"></td>
        </tr>
		 <tr>
          <td height="22" bgcolor="#FFFFFF"><div align="right">������</div></td>
          <td height="22" bgcolor="#FFFFFF">&nbsp;<span class="word_grey">
            <select name="selected">
              <%if rs.eof and rs.bof then%>
			  <option value="0" selected>--���ް���--</option>
			     <%else%>
				 <option selected value="0">-��ѡ�����-</option>
				<%For i=1 to rs.recordcount%>
				 <option value="<%=rs("ID")%>"><%=rs("UserName")%></option>
				 <%
					rs.movenext
					Next				
			  end if
			  %>
          </select>
          </span></td>
        </tr>
		
		
		<tr>
          <td height="22" colspan="2" bgcolor="#FFFFFF"><div align="center">
            <input name="Submit3" type="button" class="btn_grey" value="����" onClick="mycheck();">
            &nbsp;&nbsp;
            <input type="reset" class="btn_grey" value="��д">
          </div></td>
        </tr>
  </form>
</table>
</body>
</html>
