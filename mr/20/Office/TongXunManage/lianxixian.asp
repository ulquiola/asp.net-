<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file=../conn/conn.asp-->
<%
      set rs2=server.CreateObject("adodb.recordset")
	  sql2="select * from Tab_tongxunadd where id="&request("id")
	  rs2.open sql2,conn,1,1
%>
<%%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
<style type="text/css">
<!--
.style1 {
	font-size: 16px;
	font-weight: bold;
	color: #0000FF;
}
body,td,th {
	font-size: 12px;
}
a:link {
	color: #0000FF;
}
a:visited {
	color: #0000FF;
}
a:hover {
	color: #0000FF;
}
.style2 {color: #FFFFFF}
.style3 {
	font-size: 14px;
	color: #000000;
}
.STYLE4 {
	font-size: 10pt;
	color: #FFFFFF;
}
-->
</style>
</head>

<body>
<form name="form1" method="post" action="">
  <table width="337" height="478" border="0" align="center" cellspacing="1" bgcolor="#6DBAFF">
    <tr>
      <td height="27" colspan="2"><div align="center" class="STYLE4">ͨѶ����ϸ��Ϣ</div></td>
    </tr>
    <tr>
      <td width="74" bgcolor="#FFFFFF"><div align="right">ͨѶ������</div></td>
      <td width="256" bgcolor="#FFFFFF">
	  <%
	  set rs=server.CreateObject("adodb.recordset")
	  sql="select * from Tab_tongxun where id="&rs2("name1")
	  rs.open sql,conn,1,3
	 %>
	 <%=rs("name1")%>      </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">��&nbsp;&nbsp;&nbsp;&nbsp;����</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("name11")%></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">�������ڣ�</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("birthday")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">��&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("sex")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">����״����</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("hy")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">������λ��</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("dw")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">�������ţ�</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("department")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">ְ&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("zw")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">����ʡ�ݣ�</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("sf")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">���ڳ��У�</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("cs")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">�칫�绰��</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("phone")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">�ƶ��绰��</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("phone1")%></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">��ͥ��ַ��</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("address")%></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">OICQ��</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("OICQ")%></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">��ͥ�绰��</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("family")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">�������룺</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("postcode")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">email:</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("email")%></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">��&nbsp;&nbsp;&nbsp;&nbsp;ע��</div></td>
      <td bgcolor="#FFFFFF"><textarea name="textarea" cols="30" rows="4"><%=rs2("remark")%>

</textarea></td>
    </tr>
    <tr bgcolor="#F3F3F3">
      <td colspan="2"><div align="center"><a href="#"><img src="../Images/waichuan8.GIF" width="62" height="31" border="0" onClick="JScript:window.close();"></a>
      </div></td>
    </tr>
  </table>
</form>
</body>
</html>
<%session("id")=request("id")%>
