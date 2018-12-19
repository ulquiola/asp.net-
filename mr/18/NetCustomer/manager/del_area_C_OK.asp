<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<%
On error resume next  '设置错误陷阱
If (Request.Form("City") <> "") Then 
  vCity = Request.Form("City")
End If
'查询用户所输入的省级名称是否包括下属地区
Set rs_father = Server.CreateObject("ADODB.Recordset")
sql_F = "SELECT * FROM DB_Area WHERE Father = " + Replace(vCity, "'", "''") + ""
rs_father.open sql_F,conn,1,3
if not rs_father.eof then%>
<script language="javascript">
alert("该市/县还包括下属地区，请先删除其下属地区！");
window.close();
</script>
<%else
'判断用户输入的省份是否包括客户信息
sql="delete from DB_area where ID="&request.Form("city")
set rs=conn.execute(sql)
if err then%>
<script language="javascript">
alert("该区域已包括客户信息，数据删除失败！");
opener.location.reload();  //刷新其父窗口
window.close();
</script>
<% end if %>
<script language="javascript">
alert("数据删除成功！");
opener.location.reload();
window.close();
</script>
<%end if%>
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
            <form name="form1" method="post">
            
            <table width="80%" height="147" border="1" cellpadding="-2"  cellspacing="-2"
			 bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr>
                <td height="97" bgcolor="#809EA4"><div align="center" class="style1">数据删除成功！</div>
                  <div align="left">
</div></td>
                </tr>
              <tr>
                <td height="33"><div align="center">
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