<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->

<%
'查询管理员信息
if request("ID")<>"" then
session("ID")=request("ID")
end if
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM DB_manager WHERE ID = "&session("ID")&""
rs.open sql,conn,1,3%>
<%
'删除管理员信息
if request.Form("Name")<>"" then
	Del="Delete from DB_manager WHERE ID = "&session("ID")&""
	conn.execute(Del)%>
	  <script language="javascript">
	  alert("数据删除成功！");
	  opener.location.reload();
	  window.close();
	  </script>
<%end if%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>管理员登录--修改管理员信息！</title>
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
<table width="400" height="242" border="0" cellpadding="-2" cellspacing="-2" background="../Images/add_manager.gif">
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
            <form ACTION="del_manager.asp" METHOD="POST" name="form1">
            
            <table width="80%" height="136" border="1" cellpadding="-2"  cellspacing="-2" bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr>
                <td width="28%" bgcolor="#809EA4"><div align="center" class="style1">管理员名称：</div></td>
                <td width="72%">
                  <div align="left">
                    &nbsp;
                    <input name="Name" type="text" class="Sytle_auto" id="Name" value="<%=rs("Name")%>" readonly="yes">
                  </div></td>
                </tr>
              <tr>
                <td bgcolor="#809EA4"><div align="center"><span class="style1">&nbsp;管理员密码：</span></div></td>
                <td><div align="left">&nbsp;
                    <input name="PWD" type="password" class="Sytle_auto" id="PWD" value="<%=rs("PWD")%>">
</div></td>
                </tr>
              <tr>
                <td bgcolor="#809EA4" class="style1"><div align="center">管理员权限：</div></td>
                <td><div align="left">&nbsp;
                  <select name="purview" id="purview">
                    <option value="系统" <%If (Not isNull(rs("purview"))) Then If ("系统" = CStr(rs("purview"))) Then Response.Write("SELECTED") : Response.Write("")%>>系统　　　</option>
                    <option value="一般" <%If (Not isNull(rs("purview"))) Then If ("一般" = CStr(rs("purview"))) Then Response.Write("SELECTED") : Response.Write("")%>>一般　　　</option>
                  </select>
                </div></td>
              </tr>
              <tr>
                <td colspan="2"><div align="center">
                  <input name="Submit" type="submit" class="Style_button" value="删除">
                  <input name="Button" type="button" class="Style_button" value="关闭" onClick="javascrip:opener.location.reload();window.close()">
                </div></td>
                </tr>
              </table>
            </form>
            </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
