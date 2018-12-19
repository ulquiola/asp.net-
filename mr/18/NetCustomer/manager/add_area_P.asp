<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<%
if request.Form("name")<>"" then
	vname=request.Form("name")
	vtype=request.Form("Type")
	'查询用户输入的省级名称
	set rs=server.CreateObject("ADODB.RecordSet")
	sql_rs="select * from DB_Area where name='"&vname&"' and typeID="&vtype
	rs.open sql_rs,conn,1,3
	if rs.eof then
	'在地域信息表中插入该省级名称
		sql="Insert Into DB_Area (name,typeID) values('"&vname&"',"&vtype&")"
		conn.execute(sql)%>
		<script language="javascript">
		alert("数据添加成功！");
		opener.location.reload(); //刷新其父窗口
		window.close();　　//关闭当前窗口
		</script>
	<% else %>
		<script language="javascript">
		alert("该地域信息已经存在！");
		window.close();
		</script>
	<% end if
end if %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>管理员登录--添加地域信息！</title>
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
<script language="javascript">
function mycheck(){
if (form1.Name.value=="")
{
type=""
switch(form1.type.value)
{
case "1":type="省";break;
case "4":type="直辖市";break;
case "5":type="自治区";break;
}
alert("请输入 ["+type+"] 名称！");form1.Name.focus();return;}
form1.submit();
}
</script>
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
            <form ACTION="add_area_P.asp" METHOD="POST" name="form1">
            
            <table width="80%" border="1" cellpadding="-2"  cellspacing="-2" bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr>
                <td width="31%" height="36" bgcolor="#809EA4"><div align="center" class="style1">省级名称：</div></td>
                <td width="69%" height="36">
                  <div align="left">
                    &nbsp;
                    <input name="Name" type="text" class="Sytle_auto" id="Name">
                  </div></td>
                </tr>
              <tr>
                <td height="36" bgcolor="#809EA4"><div align="center"><span class="style1">&nbsp;区域类型：</span></div></td>
                <td height="36"><div align="left">&nbsp;
                    <select name="type" id="type">
                      <option value="1" selected>省</option>
                      <option value="4">市</option>
                      <option value="5">自治区</option>
                    </select>
</div></td>
                </tr>
              <tr>
                <td height="46" colspan="2"><div align="center">
                  <input name="Button" type="button" class="Style_button" value="保存" onClick="mycheck()">
                  <input name="Submit2" type="reset" class="Style_button" value="重置">
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
