<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<%
if request.Form("Name")<>"" then
	Father=request.Form("father")
	typeID=request.Form("TypeID")
	vName=request.Form("Name")
	set rs=server.CreateObject("ADODB.RecordSet")
	sql_rs="select * from DB_Area where name='"&vname&"' and typeID="&typeID&" and father="&father
	rs.open sql_rs,conn,1,3
	if rs.eof then
	sql="Insert Into DB_area (father,typeid,name) values("&father&","&typeID&",'"&vName&"')"
	conn.execute (sql)%>
	<script language="javascript">
	alert("数据添加成功！");
	opener.location.reload();
	window.close();
	</script>
	<% else %>
	<script language="javascript">
	alert("该地域信息已经存在！");
	window.close();
	</script>
	<% end if
end if %>
<%
Set rs_province = Server.CreateObject("ADODB.Recordset")
sql_P = "SELECT *  FROM DB_Area  WHERE TypeID = 1"
rs_province.open sql_P,conn,1,3
%>
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
{alert("请输入 [市/县] 名称！");form1.Name.focus();return;}
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
            <form ACTION="add_area_C.asp" METHOD="POST" name="form1">
            
            <table width="80%" border="1" cellpadding="-2"  cellspacing="-2" bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr>
                <td width="31%" height="36" bgcolor="#809EA4"><div align="center" class="style1"> 所属省份：</div></td>
                <td width="69%" height="36">
                  <div align="left">
&nbsp; 
<select name="Father" id="select">
  <%
While NOT rs_province.EOF
%>
  <option value="<%=rs_province("ID")%>"><%=rs_province("Name")%></option>
  <%
  rs_province.MoveNext()
Wend
%>
</select>
                  <input name="TypeID" type="hidden" id="TypeID" value="2">
                  </div></td>
                </tr>
              <tr>
                <td height="36" bgcolor="#809EA4"><div align="center"><span class="style1">市/县级名称&nbsp;：</span></div></td>
                <td height="36"><div align="left">&nbsp; 
				<input name="Name" type="text" class="Sytle_auto" id="Name">
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
