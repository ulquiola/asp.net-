<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<%
Set rs_province = Server.CreateObject("ADODB.Recordset")
sql_P = "SELECT *  FROM DB_Area  WHERE TypeID=1 or TypeID=4 or TypeID=5"
rs_province.open sql_P,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>管理员登录--删除地域信息！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image:  url(../Images/bg.gif);
}
.style1 {color: #FFFFFF}
.style2 {color: #a2bcc5}
-->
</style></head>
<body>
<table width="400" height="242" border="0" cellpadding="-2" cellspacing="-2"
 background="../Images/add_manager.gif">
  <tr>
    <td valign="top"><table width="400" height="271" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="11" height="85">&nbsp;</td>
        <td width="373">&nbsp;</td>
        <td width="14"><span class="style2"></span></td>
      </tr>
      <tr>
        <td height="144">&nbsp;</td>
        <td valign="top">
  
          <div align="center">
            <form name="form1" method="post" action="del_area_P_OK.asp">
            
            <table width="80%" height="141" border="1" cellpadding="-2"  cellspacing="-2"
			 bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr>
                <td height="86">请选择要删除的省级名称[包括自治区、直辖市]：<br>                  <br>
 <select name="Province" id="select4">
    <%While (NOT rs_province.EOF)%>
    <option value="<%=rs_province("ID")%>"><%=rs_province("Name")%></option>
    <% rs_province.MoveNext()
	Wend %>
  </select>                  </td>
                </tr>
              <tr>
                <td height="53"><div align="center">
                  <input name="Submit" type="submit" class="Style_button" value="删除">
                  <input name="Button" type="button" class="Style_button" value="关闭"
				   onClick="javascrip:opener.location.reload();window.close()">
</div></td>
              </tr>
              </table>
            </form>
            </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td></td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>