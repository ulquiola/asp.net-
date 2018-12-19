<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp" -->
<%
Set rs_personnel = Server.CreateObject("ADODB.Recordset")
sql_P= "SELECT * FROM Tab_User WHERE ID = "&Request.QueryString("ID")&""
rs_personnel.open sql_P,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>员工详细信息！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
.style1 {color: #C60001}
.style2 {color: #669999}
.STYLE4 {
	font-size: 9pt;
	color: #FFFFFF;
}
-->
</style></head>

<body>
<table width="557" height="401" border="0" cellpadding="-2" cellspacing="-2" background="../Images/detail0.gif">
  <tr>
    <td width="817" height="349" valign="top"><table width="100%" height="94"  border="0" cellpadding="-2" cellspacing="-2">
      <tr>
        <td height="43" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE4">员工详细信息</span></td>
      </tr>
      <tr>
        <td height="51">&nbsp;</td>
      </tr>
    </table>      
      <form name="form1">
        <table width="83%" height="214"  border="0" align="center" cellpadding="-2" cellspacing="-2">
          <tr>
            <td width="14%" height="27"><div align="center" class="style1">用户名：</div></td>
            <td width="32%"><input name="username" type="text" class="Sytle_text" id="username" value="<%=(rs_personnel("UserName"))%>"></td>
            <td width="10%" class="style1"><div align="center">姓名：</div></td>
            <td width="44%">
            <input name="name" type="text" class="Sytle_text" id="name" value="<%=(rs_personnel("Name"))%>">            </td>
          </tr>
          <tr>
            <td height="27" valign="top"><div align="center" class="style1">部门：</div></td>
            <td><input name="branch" type="text" class="Sytle_text" id="branch" value="<%=rs_personnel("branch")%>"></td>
            <td class="style1"><div align="center">权限：</div></td>
            <td>&nbsp;<%=rs_personnel("purview")%>　　<span class="style1">性别：</span><%=(rs_personnel("sex"))%></td>
          </tr>
          <tr>
            <td height="27" valign="top"><div align="center" class="style1">电话：</div></td>
            <td><input name="tel" type="text" class="Sytle_text" id="tel" value="<%=(rs_personnel("Tel"))%>"></td>
            <td class="style1"><div align="center">职务：</div></td>
            <td><input name="job" type="text" class="Sytle_text" id="job" value="<%=(rs_personnel("job"))%>"></td>
          </tr>
          <tr>
            <td height="27" valign="top" class="style1"><div align="center">Email：</div></td>
            <td colspan="3"><input name="email" type="text" class="Style_subject" id="email" value="<%=(rs_personnel("Email"))%>"></td>
          </tr>
          <tr>
            <td height="27" valign="top" class="style1"><div align="center">地址：</div></td>
            <td colspan="3"><input name="address" type="text" class="Style_subject" id="address" value="<%=(rs_personnel("Address"))%>"></td>
          </tr>
          <tr>
            <td colspan="4"><div align="center">
                <br>
                <input name="myclose" type="button" class="Style_button_del" id="myclose" value="关闭窗口" onClick="javascrip:window.close()">
                </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>
