<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"--><!--�������ݿ������ļ�-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")			'������¼��
sql="Select * From tb_User Where Status<>'����Ա' order by id desc"		'��ѯ����
rs.Open sql,conn,1,3									'�򿪼�¼��
%>
<html>
<head>
<title>�û�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Css/style.css" rel="stylesheet">
<script src="JS/fun.js"></script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.STYLE1 {
	font-size: 10pt;
	font-weight: bold;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
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
<%if rs.eof then
response.Write("�����û���Ϣ")
else
%>
      <table width="100%" height="205"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder">
        <tr>
          <td width="13" height="32" bgcolor="#4EAEE1">&nbsp;</td>
          <td width="689" bgcolor="#4EAEE1" style="color:#FFFFFF"> ��<span class="STYLE1"><font color="#FFFFFF">��ǰλ�ã��鿴�û���Ϣ</font></span> �� </td>
          <td width="73" bgcolor="#4EAEE1">&nbsp;</td>
        </tr>
        <tr valign="top">
          <td height="171" colspan="3"><table width="100%" height="170"  border="0" cellpadding="-2" cellspacing="-2">
            <tr>
              <td height="10" valign="top"></td>
            </tr>
            <tr>
              <td valign="top"><table width="100%" height="74"  border="2" cellpadding="0" cellspacing="0" bordercolor="#4EAEE1" bordercolorlight="#FFFFFF" bgcolor="#FFFFFF" class="tableBorder_B">
                <tr align="center">
                  <td width="112">�û���</td>
                  <td width="102" height="25">��ʵ����</td>
                  <td width="32">�Ա�</td>
                  <td width="108">����</td>
                  <td width="210">Email��ַ</td>
                  <td width="78">QQ</td>
                  <td width="56">��������</td>
                  <td colspan="2">����</td>
                </tr>
		<%
		'��ҳ��ʾ�û���Ϣ
		rs.pagesize=15							'���÷�ҳ��ʾʱ��ÿҳ��ʾ��¼����Ŀ
		page1=CLng(Request("page1"))				'����ȡ���ļ�¼��ת��Ϊ����
		if page1<1 then page1=1						'Ϊpage��������Ĭ��ֵ
		rs.absolutepage=page1						'�ڽ��з�ҳ��ʾʱ��ָ��ָ�����ڵ�ҳ
		for i=1 to rs.pagesize						'ѭ����ʾ��¼��Ϣ
		ID=rs("ID")									'��ȡ��¼����ID��ֵ������ID����
		%>
					<script language="javascript">
				  	function del(ID){				//����del�Զ��庯��
					if(confirm("���Ҫɾ�����û���")){	//������ʾ�Ի���
						window.location.href="User_del.asp?UID="+ID;//��ת��ָ����ҳ��
					 }
					}
				  </script>
                <tr>
                  <td align="center"><%=rs("UserName")%></td>
                  <td height="25" align="center"><%=rs("TrueName")%></td>
                  <td align="center"><%=rs("Sex")%></td>
                  <td align="center"><%=rs("Birthday")%></td>
                  <td align="center"><%=rs("Email")%></td>
                  <td align="center"><%=rs("QQ")%>&nbsp;</td>
                  <td align="center"><%=rs("SendRatio")%></td>
                  <td width="37" align="center"><a href="#" onClick="del(<%=ID%>)"><img src="../images/del.gif" width="22" height="22" border="0" /></a></td>
                  <td width="20" align="center"><a href="modify_statc.asp?id=<%=ID%>"><img src="../Images/modify.gif" width="20" height="18" border="0"></a></td>
                </tr>
		<%
			  rs.MoveNext					'�����ƶ���¼ָ��
			  If rs.Eof Then exit For 		'�˳�Forѭ��
		      Next
		%> 
              </table>
                <table width="100%" height="26"  border="0" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
	<%If not(rs.Eof and rs.Bof) Then%>	
    <td width="34%" align="left">&nbsp;
	      		<% if page1<>1 then %><a href=?page1=1 class="white">��һҳ</a>
				<a href=?page1=<%=(page1-1)%> class="white">��һҳ</a> 
			<%
			end if 
			if page1<>rs.pagecount then %>
   				<a href=?page1=<%=(page1+1)%> class="white">��һҳ</a> 
				<a href=?page1=<%=rs.pagecount%> class="white">���һҳ</a> 
		  <%end if%>
		  </td>
    <td width="66%" align="right" class="word_grey">[<%=page1%>/<%=rs.PageCount%>]&nbsp;&nbsp;ÿҳ<%=rs.PageSize%>��&nbsp;&nbsp;��<%=rs.RecordCount%>λ�û�&nbsp;</td>
		<%End If%>	
                </table></td>
            </tr>
            <tr>
              <td height="10" valign="top"></td>
            </tr>
          </table></td>
        </tr>
      </table>
<%end if%>
    </td>
  </tr>
</table>
</body>
</html>
<%
rs.close()				'�رռ�¼��
Set rs=Nothing			'�ͷ�Recordset����ռ�õ�������Դ
%>