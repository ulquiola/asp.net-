<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp" -->
<%
if request.Form("UserName")<>"" then
	DEL="Delete From Tab_User where ID='"&session("ID")&"'"
	conn.execute(Del)%>
	<script language="javascript">
	alert("员工信息已经删除！");
	opener.location.reload();
	window.close();
	</script>
<%end if%>
<%
if Request.QueryString("ID")<>""then
	session("ID")=Request.QueryString("ID")
end if
Set rs_personnel = Server.CreateObject("ADODB.Recordset")
sql_P="SELECT ID,UserName,Name,PWD,purview,branch,sex,Email,Tel,Address"&_
" FROM Tab_User WHERE ID="&session("ID")&""
rs_personnel.open sql_P,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>删除员工信息！</title>
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
<table width="460" height="403" border="0" cellpadding="-2" cellspacing="-2" background="../Images/waichudeng.gif">
  <tr>
    <td width="817" height="403" valign="top"><table width="100%" height="89"  border="0" cellpadding="-2" cellspacing="-2">
      <tr>
        <td height="41" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE4">删除员工信息</span></td>
      </tr>
      <tr>
        <td height="44">&nbsp;</td>
      </tr>
    </table>      
      <br>
      <form ACTION="personnel_del.asp" METHOD="POST" name="form1">
        <table width="95%" height="222"  border="0" align="center" cellpadding="-2" cellspacing="-2">
          <tr>
            <td colspan="6"><span class="style1">&nbsp;&nbsp;&nbsp;用户名：</span>&nbsp;<%=(rs_personnel("UserName"))%></td>
          </tr>
		  <tr>
            <td width="16%" height="27"><div align="center" class="style1">姓名：</div></td>
            <td width="39%" height="27"><span class="style1">
              <input name="username" type="text" class="Sytle_text" id="username2" readonly value="<%=(rs_personnel("Name"))%>">
            </span> </td>
            <td width="9%" height="27" class="style1">权限：</td>
            <td width="13%" height="27"><select name="purview" id="purview">
              <option value="系统" <%If (Not isNull((rs_personnel("purview")))) Then If ("系统" = CStr((rs_personnel("purview")))) Then Response.Write("SELECTED") : Response.Write("")%>>系统</option>
              <option value="读写" <%If (Not isNull((rs_personnel("purview")))) Then If ("读写" = CStr((rs_personnel("purview")))) Then Response.Write("SELECTED") : Response.Write("")%>>读写</option>
              <option value="只读" <%If (Not isNull((rs_personnel("purview")))) Then If ("只读" = CStr((rs_personnel("purview")))) Then Response.Write("SELECTED") : Response.Write("")%>>只读</option>
            </select></td>
            <td width="10%" height="27"><span class="style1">性别：</span></td>
            <td width="13%" height="27"><span class="style1">
              <select name="sex" id="select2">
                <option value="男" <%If (Not isNull((rs_personnel("sex")))) Then If ("男" = CStr((rs_personnel("sex")))) Then Response.Write("SELECTED") : Response.Write("")%>>男</option>
                <option value="女" <%If (Not isNull((rs_personnel("sex")))) Then If ("女" = CStr((rs_personnel("sex")))) Then Response.Write("SELECTED") : Response.Write("")%>>女</option>
              </select>
            </span></td>
          </tr>
          <tr>
            <td height="27"><div align="center" class="style1">电话：</div></td>
            <td height="27"><input name="tel" type="text" class="Sytle_text" id="tel" value="<%=(rs_personnel("Tel"))%>"></td>
            <td height="27" class="style1">部门：</td>
            <td height="27" colspan="3"><select name="branch" id="branch">
              <option value="开发部" <%If (Not isNull((rs_personnel("branch")))) Then If ("开发部" = CStr((rs_personnel("branch")))) Then Response.Write("SELECTED") : Response.Write("")%>>开发部</option>
              <option value="人事部" <%If (Not isNull((rs_personnel("branch")))) Then If ("人事部" = CStr((rs_personnel("branch")))) Then Response.Write("SELECTED") : Response.Write("")%>>人事部</option>
              <option value="销售部" <%If (Not isNull((rs_personnel("branch")))) Then If ("销售部" = CStr((rs_personnel("branch")))) Then Response.Write("SELECTED") : Response.Write("")%>>销售部</option>
            </select></td>
          </tr>
          <tr>
            <td height="27" class="style1"><div align="center">Email：</div></td>
            <td height="27" colspan="5"><input name="email" type="text" class="Style_subject" id="email" value="<%=(rs_personnel("Email"))%>"></td>
          </tr>
          <tr>
            <td height="27" class="style1"><div align="center">地址：</div></td>
            <td height="27" colspan="5"><input name="address" type="text" class="Style_subject" id="address" value="<%=(rs_personnel("Address"))%>"></td>
          </tr>
          <tr>
            <td colspan="6"><div align="center">
                <br>
                <input name="Button" type="submit" class="Style_button_del" value="确定删除">
                <input name="Submit2" type="reset" class="Style_button" value="重置">
                <input name="myclose" type="button" class="Style_button" id="myclose" value="关闭" onClick="javascrip:window.close()">
            </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>