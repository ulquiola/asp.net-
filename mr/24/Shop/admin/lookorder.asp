<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="chk_admin.asp" -->
<!--#include file="../include/include.asp" -->
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
<%
if request("dingdan")<>"" and request("action")="update" then	'�ж��Ƿ��޸�
	sql="select * from [dingdan] where didanhao="&request("dingdan")&";"	'�������Ų�ѯ
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	if not rs.eof then
		if right(request("zhuangtai"),1)>="2" and rs("zhuangtai")<"2" then
		'��ʱ request("zhuangtai") Ϊ�ַ�����ʽ������1,2��1,2,3 ���ʱ������ֻҪ���ұߵ��ַ��Ϳ���֪���ύ��ֵ��
		'�� right(request("zhuangtai"),1) ��ֵ�������� 2 ���Ѿ������ˣ��������ݿ��е�ֵС�� 2 (��Ʒ����û�б��޸Ĺ�)��ʱ��Ŷ���Ʒ�����������޸�
			shangpin=split(rs("shangpin"),",")
			shuliang=split(rs("shuliang"),",")
			for i=0 to ubound(shangpin)	'ѭ�������Ʒ ID���ж�����Ʒ�Ͷ�Ӧ��Ʒ ID���������޸�
				sql2="select * from [shangpin] where id="&shangpin(i)&""
				set rs2=Server.CreateObject("ADODB.Recordset")
				rs2.open sql2,conn,3,3
				rs2("shuliang")=rs2("shuliang")-shuliang(i)	'�µ���Ʒ����=ԭ��Ʒ����-�����е���Ʒ����
				rs2.update
				rs2.close
				set rs2=nothing
			next
		end if
		if request("zhuangtai")<>"" then
			rs("zhuangtai")=right(request("zhuangtai"),1)
		else
			rs("zhuangtai")=0	'�������״̬���������տ�ѷ��������ջ��������
		end if
		rs.update
		rs.close
		set rs=nothing
		response.Write("<script>alert('�޸Ķ����ɹ���');window.location.href='lookorder.asp?dingdan="&request("dingdan")&"';</script>")
	end if
end if
if request("dingdan")<>"" then
	SafeRequest(request("dingdan"))	'�ж϶������Ƿ�Ϊ������
	sql="select * from [dingdan] where didanhao="&request("dingdan")&";"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	if not rs.eof then
		shangpin=split(rs("shangpin"),",")	'�����ָ��ַ���������Ҫ��ִ˶�����ÿ����Ʒ�� ID
		shuliang=split(rs("shuliang"),",")	'��Ʒ��Ӧ�Ĺ�������
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
    <tr>
      <td align="center"><font color="#FFFFFF">��Ʒ��������</font></td>
    </tr>
    <tr>
      <td valign="top" bgcolor="#FFFFFF"><br>
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td><table border="0" cellspacing="1" cellpadding="4" align="center" width="100��" bgcolor="#6699FF">
	<form action="lookorder.asp" method="get">
      <tr>
        <td width="13%" bgcolor="#FFFFFF">�������:</td>
        <td width="22%" bgcolor="#FFFFFF"><%=rs("didanhao")%></td>
        <td width="18%" bgcolor="#FFFFFF"><div align="center">���տ�
              <input type="checkbox" name="zhuangtai" value="1"<%if rs("zhuangtai")>0 then response.Write("checked") end if	'ȷ����Ʒ״̬%>>
        </div></td>
        <td width="18%" bgcolor="#FFFFFF"><div align="center">�ѷ���
              <input type="checkbox" name="zhuangtai" value="2"<%if rs("zhuangtai")>1 then response.Write("checked") end if%>>
        </div></td>
        <td width="18%" bgcolor="#FFFFFF"><div align="center">���ջ�
              <input type="checkbox" name="zhuangtai" value="3"<%if rs("zhuangtai")>2 then response.Write("checked") end if%>>
        </div></td>
        <td width="7%" bgcolor="#FFFFFF"><div align="right">
          <input type="submit" name="submit" value="�޸�">
        </div></td>
      </tr>
	  <input type="hidden" name="dingdan" value="<%response.Write rs("didanhao")	'����������Ϊ����ֵ�����ύ%>">
	  <input type="hidden" name="action" value="update">
	  </form>
    </table></td>
  </tr>
  <tr>
    <td>
      <table border="0" cellspacing="1" cellpadding="4" align="center" width="100��" bgcolor="#6699FF">
        <tr bgcolor="#FFFFFF" height="25" align="center">
          <td width="300">�� Ʒ �� ��</td>
          <td width="40">����</td>
          <td width="60">�г���</td>
          <td width="60">��Ա��</td>
          <td width="60">�ɽ���</td>
          <td width="70">С ��</td>
        </tr>
        <tr bgcolor="#FFFFFF" height="25"align="center">
          <td align="left">
<%
		for i=0 to ubound(shangpin)	'��ʱ���� shangpin Ϊ������ʽ
			sql2="select * from [shangpin] where id="&trim(shangpin(i))&""	'�õ�ÿ����Ʒ�� ID
			set rs2=Server.CreateObject("ADODB.Recordset")
			rs2.open sql2,conn,3,3
%>
<a href="pro.asp?id=<%=rs2("id")%>"><%response.Write rs2("mingcheng") '�����Ʒ����%></a><br>
<%
			rs2.close
			set rs2=nothing
		next
%>
		  </td>
          <td>
<%
		for i=0 to ubound(shuliang)
			response.Write(shuliang(i)) '�����Ʒ����
			response.Write("<br>")
		next
%>
		</td>
          <td>
<%
		for i=0 to ubound(shuliang)	'��ʱ���� shuliang �� shangpin ������±���һ����
			'��Ϊ�ڴ洢����ʱ��Ʒ��Ӧ��������������� FOR ѭ�������Ҫ���ֵ�������κ�һ����������Ϊ FOR ѭ�������ֵ
			sql2="select * from [shangpin] where id="&trim(shangpin(i))&""
			set rs2=Server.CreateObject("ADODB.Recordset")
			rs2.open sql2,conn,3,3
			response.Write(rs2("shichang")) '�����Ʒ�г��۸�
			response.Write("<br>")
			rs2.close
			set rs2=nothing
		next
%>
		  </td>
          <td>
<%
		for i=0 to ubound(shuliang)
			sql2="select * from [shangpin] where id="&trim(shangpin(i))&""
			set rs2=Server.CreateObject("ADODB.Recordset")
			rs2.open sql2,conn,3,3
			response.Write(rs2("huiyuan")) '�����Ʒ��Ա�۸�
			response.Write("<br>")
			rs2.close
			set rs2=nothing
		next
%>
		  </td>
          <td>
<%
		for i=0 to ubound(shuliang)
			sql2="select * from [shangpin] where id="&trim(shangpin(i))&""
			set rs2=Server.CreateObject("ADODB.Recordset")
			rs2.open sql2,conn,3,3
			response.Write(rs2("huiyuan")) '�����Ʒ�ɽ��۸�(�ɽ��۸���ʵ���ǻ�Ա�۸�)
			response.Write("<br>")
			rs2.close
			set rs2=nothing
		next
%>
		  </td>
          <td>
<%
		for i=0 to ubound(shuliang)
			sql2="select * from [shangpin] where id="&trim(shangpin(i))&""
			set rs2=Server.CreateObject("ADODB.Recordset")
			rs2.open sql2,conn,3,3
			response.Write(rs2("huiyuan")*shuliang(i)) '���С��
			response.Write("<br>")
			rs2.close
			set rs2=nothing
		next
%>
		  </td>
        </tr>
    </table></td>
  </tr>
</table>
<table width="600"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="center" class="style1">ע��һ����Ʒȷ������������Ʒ�������Զ��ӿ������Ӧ���٣����Ҳ��ɸ��ģ�</div></td>
  </tr>
</table>
<table width="600" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6699FF">
        <tr bgcolor=#ffffff>
          <td width="150">�ջ���������</td>
          <td width="600" height="28"><%=rs("shoujianren")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td>��ϸ��ַ��</td>
          <td height="28"><%=rs("dizhi")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td>�ʡ����ࣺ</td>
          <td height="28"><%=rs("youbian")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td>�硡������</td>
          <td height="28"><%=rs("tel")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td>�����ʼ���</td>
          <td height="28"><%=rs("mail")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td>�ͻ���ʽ��</td>
          <td height="28"><%=rs("songhuo")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td>֧����ʽ��</td>
          <td height="20"><%=rs("zhifu")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td>�����ԣ�</td>
          <td height="28"><%response.Write HTMLEncode(rs("liuyan"))	'���� HTMLEncode �Ĺ��ܾ����滻�ո񡢻���,������ include.asp �ļ���%></td>
        </tr>
</table>
<%
	else
		response.Write("<script>alert('�޴˶�����');</script>")
	end if
	rs.close
	set rs=nothing
end if
%>		  <br></td>
    </tr>
</table>
