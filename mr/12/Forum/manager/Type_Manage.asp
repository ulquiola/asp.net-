<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"--><!--�������ݿ������ļ�-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")			'������¼��
SQL="Select * from view_Type"							'��ѯ����
rs.Open SQL,conn,1,3									'�򿪼�¼��
%>
<html>
<head>
<title>�༭������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Css/style.css" rel="stylesheet">
<script src="../JS/fun.js"></script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>
<body  bgcolor="B9DFFF">
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
          <td width="689" bgcolor="#4EAEE1" style="color:#FFFFFF">&nbsp;&nbsp;<font color="#FFFFFF"><strong>��ǰλ�ã��༭������</strong></font></td>
          <td width="73" bgcolor="#4EAEE1">&nbsp;</td>
        </tr>
        <tr valign="top">
          <td height="171" colspan="3"><table width="100%" height="189"  border="0" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
            
<%If rs.eof and rs.bof Then%>
            <tr>
			<td align="center">������������Ϣ��</td>
            </tr>
<%Else%>
            <tr>
              <td valign="top"><table width="100%"  border="1" cellpadding="-2" cellspacing="-2" bordercolor="#4EAEE1" bordercolorlight="#FFFFFF" class="tableBorder_B">
                <tr align="center">
                  <td width="326" height="20">����������</td>
                  <td width="204">����</td>
                  <td width="301">���ʱ��</td>
                  <td width="82">�޸�</td>
                  <td width="85">ɾ��</td>
                </tr>
		<%
		'��ҳ
		rs.pagesize=15							'���÷�ҳ��ʾʱ��ÿҳ��ʾ��¼����Ŀ
		page=CLng(Request("page"))				'����ȡ���ļ�¼��ת��Ϊ����
		if page<1 then page=1					'Ϊpage��������Ĭ��ֵ
		rs.absolutepage=page					'�ڽ��з�ҳ��ʾʱ��ָ��ָ�����ڵ�ҳ
		for i=1 to rs.pagesize					'ѭ����ʾ��¼��Ϣ
		ID=rs("ID")								'��ȡ��¼����ID��ֵ������ID����
		%>
					<script language="javascript">
				  	function del(ID){									//����del�Զ��庯��
					if(confirm("���Ҫɾ���������")){				//������ʾ�Ի���
						window.location.href="Type_del.asp?ID="+ID;		//��ת��ָ����ҳ��
					 }
					}
				  </script>
                <tr>
                  <td align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[ <%=rs("TName")%> ]</td>
                  <td align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("owner")%></td>
                  <td align="left"><div align="center"><%=rs("CreateTime")%>&nbsp;</div></td>
                  <td align="center"><a href="#" onClick="window.open('modify_Type.asp?TypeID=<%=rs("ID")%>','','width=240,height=139')"><img src="../Images/modify.gif" width="20" height="18" border="0"></a></td>
                  <td align="center"><input type="image" src="../Images/del.gif" width="22" height="18" border="0" onClick="del(<%=ID%>)" class="noborder"></td>
                </tr>
		<%
			  rs.MoveNext					'�����ƶ���¼ָ��
			  If rs.Eof Then exit For 		'�˳�Forѭ��
		  Next
		%>
              </table>
                <table width="100%"  border="0" cellspacing="-2" cellpadding="-2">
	<%If not(rs.Eof and rs.Bof) Then%>	
    <td width="34%" height="27" align="left">&nbsp;
	      		<% if page<>1 then %><a href=?page=1 class="white">��һҳ</a>
				<a href=?page=<%=(page-1)%> class="white">��һҳ</a> 
			<%end if 
			if page<>rs.pagecount then %>
   				<a href=?page=<%=(page+1)%> class="white">��һҳ</a> 
				<a href=?page=<%=rs.pagecount%> class="white">���һҳ</a> 
		  <%end if%>	</td>
    <td width="66%" align="right" class="word_grey">[<%=page%>/<%=rs.PageCount%>]&nbsp;&nbsp;ÿҳ<%=rs.PageSize%>��&nbsp;&nbsp;��<%=rs.RecordCount%>��������&nbsp;</td>
		<%End If   
	%>	
                </table></td>
            </tr>
            <tr>
              <td height="10" valign="top"></td>
            </tr>
          </table></td>
<%end If%>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
<%
rs.close()
Set rs=Nothing
%>
