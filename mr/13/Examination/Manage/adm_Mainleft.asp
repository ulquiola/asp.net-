<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������ϵͳ�˵�</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
td{font-size:9pt}
a{text-decoration:none;color:black;}
a:hover{text-decoration:underline;color:blue;}
.table_bolder{border-bottom:1px solid #000000; border-left:1px solid #000000; border-right:1px solid #000000}
-->
</style>
<script language="javascript" src="js/js.js"></script>
</head>
<body>
<table width="183" border="0" align="left" cellpadding="0" cellspacing="0">
      <tr>
        <td><img src="images/lefttop.jpg" width="183" height="48"></td>
      </tr>
  <tr>
    <td>
	<table width="183" border="0" align="left" cellpadding="0" cellspacing="0" class="table_bolder">
  <tr>
    <td height="35" align="center"><a href="admin/adm_Admin.asp" target="mainr">����Ա��Ϣ����</a></td>
  </tr>
  <tr>
    <td height="35" align="center"><a href="../Register/Manage/adm_Register.asp" target="mainr">�û���Ϣ����</a></td>
  </tr>
  <tr>
    <td height="35" align="center"><a href="adm_Mainright.asp" target="mainr">�����ɼ�����</a></td>
  </tr>
  <tr>
    <td height="35" align="center"><a href="adm_affiche.asp" target="mainr">������Ϣ����</a></td>
  </tr>
  <tr>
    <td height="35" align="center"><a href="liuyan.asp" target="mainr">���԰���Ϣ����</a></td>
  </tr>
  <tr>
    <td height="35" align="center"><a href="tre_/ad_file1.asp" target="_parent" onClick="javascript:return confirm('ȷ��Ҫ������')">���������ݿ�</a> </td>
  </tr>
  <tr>
    <td height="35" align="center"><a href="tre_/tre_jihuo.asp" target="_parent" onClick="javascript:return confirm('ȷ��Ҫ����У������')">����У����</a> </td>
  </tr>
  <tr>
    <td height="17" align="center"><a href="adm_Logout.asp" target="_parent" onClick="javascript:return confirm('ȷ��Ҫ�˳���')">
    <%
	set rs=server.CreateObject("Adodb.recordset")
	sql="select  Count(*) as counts from tb_Student where format(jointime,'yyyy-mm-dd')=format(now,'yyyy-mm-dd')"
	rs.open sql,conn,1,3
	%>
</a>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC" >
        <%do while not rs.eof%>
        <tr align="center" bgcolor="#FFFFFF">
          <td height="35"><font color="#FF0000">��������������Ϊ��<%=rs("counts")%></font></td>
        </tr>
        <% rs.movenext
 			loop
	  %>
      </table>      </td>
  </tr>
  <tr>
    <td height="35" align="center"><a href="adm_Logout.asp" target="_parent" onClick="javascript:return confirm('ȷ��Ҫ�˳���')">�˳�����</a></td>
  </tr>
</table>
	</td>
  </tr>
</table>

</body>
</html>
