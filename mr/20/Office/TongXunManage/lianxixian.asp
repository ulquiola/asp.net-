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
<title>无标题文档</title>
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
      <td height="27" colspan="2"><div align="center" class="STYLE4">通讯组详细信息</div></td>
    </tr>
    <tr>
      <td width="74" bgcolor="#FFFFFF"><div align="right">通讯组名：</div></td>
      <td width="256" bgcolor="#FFFFFF">
	  <%
	  set rs=server.CreateObject("adodb.recordset")
	  sql="select * from Tab_tongxun where id="&rs2("name1")
	  rs.open sql,conn,1,3
	 %>
	 <%=rs("name1")%>      </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">姓&nbsp;&nbsp;&nbsp;&nbsp;名：</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("name11")%></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">出生日期：</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("birthday")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">性&nbsp;&nbsp;&nbsp;&nbsp;别：</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("sex")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">婚姻状况：</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("hy")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">所属单位：</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("dw")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">所属部门：</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("department")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">职&nbsp;&nbsp;&nbsp;&nbsp;务：</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("zw")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">所在省份：</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("sf")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">所在城市：</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("cs")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">办公电话：</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("phone")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">移动电话：</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("phone1")%></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">家庭地址：</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("address")%></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">OICQ：</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("OICQ")%></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">家庭电话：</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("family")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">邮政编码：</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("postcode")%>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">email:</div></td>
      <td bgcolor="#FFFFFF"><%=rs2("email")%></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">备&nbsp;&nbsp;&nbsp;&nbsp;注：</div></td>
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
