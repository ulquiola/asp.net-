<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"--><!--�������ݿ������ļ�-->
<%
if request.Form("smalltypename")<>"" Then								'�ж�smalltypename�Ƿ�Ϊ��
	typeid=request.Form("taolun")										'��ȡ��Ԫ��taolun��ֵ������taolun����
	smalltypename=request.Form("smalltypename")							'��ȡ��Ԫ��smalltypename��ֵ������smalltypename����
	userid=request.Form("selected")										'��ȡ��Ԫ��selected��ֵ������selected����
	account=request.Form("account")										'��ȡ��Ԫ��account��ֵ������account����
	if smalltypename<>"" Then											'�ж�smalltypename�Ƿ�Ϊ��
		set rs=Server.CreateObject("ADODB.RecordSet")					'������¼��
		sql="Select * From tb_smalltype where smalltypename='"&smalltypename&"'"	'��ѯ����
		rs.open sql,conn,1,3											'�򿪼�¼��
		if not(rs.eof and rs.bof) Then									'�ж��Ƿ��м�¼
			response.Write("<script language='javascript'>alert('�ð��������Ѿ����ڣ����������룡');window.location.href='addsmall_type.asp';</script>")
			'������ʾ�Ի���
			response.End()
			'�����������Խű������в���������ظ������
		Else
			sql="Insert Into tb_smalltype (typeid,smalltypename,userid,account) values("&typeid&",'"&smalltypename&"',"&userid&",'"&account&"')"
			'Ӧ��Insert Intoʵ��������Ϣ�����
			conn.execute(sql)
			response.Write("<script language='javascript'>alert('��Ϣ��ӳɹ���');window.location.href='addsmall_type.asp';</script>")
			'������ʾ�Ի���
		End If
	End If
Else
	set rs=Server.CreateObject("ADODB.RecordSet")			'������¼��
	sql="Select * From tb_user where Status='����'"			'��ѯ����
	rs.open sql,conn,1,3									'�򿪼�¼��
	set rs_1=Server.CreateObject("ADODB.RecordSet")			'������¼��
	sql_1="Select * From tb_Type"							'��ѯ����
	rs_1.open sql_1,conn,1,3								'�򿪼�¼��
End If
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Css/style.css" rel="stylesheet">
</head>
<script language="javascript">
function mycheck()//�����Զ��庯��
{
if(form1.taolun.value==0)														//�ж��Ƿ�ѡ����������
{alert("��ѡ��������������");form1.taolun.focus();return false}				//������ʾ�Ի���
if(form1.smalltypename.value=="")												//�жϰ�������Ƿ��
{alert("�����������ƣ�");form1.smalltypename.focus();return false}			//������ʾ�Ի���
if(form1.selected.value==0)														//�ж��Ƿ�ѡ���˰���
{alert("��ѡ������������");form1.selected.focus();return false}				//������ʾ�Ի���
if(form1.account.value=="")														//�жϰ��˵���Ƿ�Ϊ��
{alert("������������˵����Ϣ��");form1.account.focus();return false}		//������ʾ�Ի���
form1.submit();																	//�ύ��
}
</script>
<body  bgcolor="B9DFFF">
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#4EAEE1">
  <form name="form1" method="post" action="addsmall_type.asp">
	    <tr>
          <td height="22" colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
            <tr>
              <td height="22" bgcolor="#4EAEE1">&nbsp;&nbsp;<font color="#FFFFFF"><strong>��ǰλ�ã���Ӱ��</strong></font></td>
            </tr>
          </table></td>
	    </tr>
        <tr>
          <td width="20%" height="10" bgcolor="#FFFFFF"><div align="right">������������</div></td>
          <td width="80%" bgcolor="#FFFFFF">
		  <select name="taolun" id="taolun">
              <%if rs_1.eof and rs_1.bof then%>
			  <option value="0" selected>--����������--</option>
			     <%else%>
				 <option selected value="0">-��ѡ��������-</option>
				<%For i=1 to rs_1.recordcount%>
				 <option value="<%=rs_1("ID")%>"><%=rs_1("TypeName")%></option>
				 <%
					rs_1.movenext
					Next				
			  end if
			  %>
          </select>
		  </td>
        </tr>
        <tr>
          <td width="20%" height="11" bgcolor="#FFFFFF"><div align="right">������ƣ�</div></td>
          <td bgcolor="#FFFFFF"><input name="smalltypename" type="text" class="inputcss1" id="smalltypename" size="50" maxlength="20"></td>
        </tr>
		 <tr>
          <td height="10" bgcolor="#FFFFFF"><div align="right">������</div></td>
          <td height="10" bgcolor="#FFFFFF"><span class="word_grey">
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
		   <td height="11" bgcolor="#FFFFFF"><div align="right">���˵����</div></td>
		   <td height="11" bgcolor="#FFFFFF"><textarea name="account" cols="40" rows="3" id="account"></textarea></td>
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
