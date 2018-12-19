<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Conn/conn.asp" -->
<%
'查询用户信息
Set rs_user = Server.CreateObject("ADODB.Recordset")
sql_user= "SELECT ID, UserName, purview FROM Tab_User WHERE UserName = '"&session("UserName")&"'"
rs_user.open sql_user,conn,1,3
'查询员工信息
Set rs_personnel = Server.CreateObject("ADODB.Recordset")
sql_P="SELECT ID,UserName,Name,purview,branch,job FROM Tab_User"&_
" ORDER BY UserName ASC"
rs_personnel.open sql_P,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>员工信息查询</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
.style10 {color: #669999}
.style11 {color: #C60001}
.STYLE13 {color: #C60001; font-size: 9pt; }
.STYLE14 {
	font-size: 9pt;
	color: #000000;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
.STYLE16 {font-size: 9pt}
-->
</style></head>
<script language="javascript">
function Mycheck()
{
if(form1.branch.value=="")
{alert("请输入员工所在部门名称??");form1.branch.focus();return;}
form1.submit();
}
</script>
<body>
<table width="814" height="505" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="488" valign="top" background="../Images/main.gif"><table width="86%" height="84%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="39" valign="bottom"><table width="270" height="20" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="29"><div align="center"><img src="../Images/isexists.gif" width="16" height="16"></div></td>
            <td width="241"><img src="../Images/y.gif" width="100" height="19"></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td>
		<table width="60%" height="21%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="69" valign="bottom"><br>
		  <table width="557" height="40"  border="0" cellpadding="-2" cellspacing="-2">
  <tr>
    <td width="7%" height="24" align="center" valign="middle">&nbsp; </td>
    <td width="20%" height="24" align="center"></td>
    <td width="4%" align="center"><img src="../images/modify.gif" width="12" height="12"></td>
    <td width="17%" align="center" class="style10"><div align="left">

	<A HREF=# class="STYLE14" onClick="Javascrip:window.open('personnel_modifyPWD.asp?ID=<%=rs_user("ID")
	 %>','','width=460,height=400')">修改个人信息</a>
	</div></td>
    <td width="5%" height="24"><div align="center" class="style10"></div></td>
    <td width="30%" height="24">    </td>
    <td width="17%" height="24"><div align="center"></div></td>
  </tr>
  <tr>
    <td height="16" colspan="8"></td>
  </tr>
</table>
<table width="556" border="0" cellspacing="-2" cellpadding="-2" align="center" >
  <tr>
    <td><form action="personnel_info.asp" method="post" name="form1">
      <table width="100%" border="0" cellspacing="-2" cellpadding="-2" align="center">
        <tr>
          <td width="2%">&nbsp;</td>
          <td width="23%">&nbsp;<span class="STYLE13">请输入员工所在部门：</span></td>
          <td width="36%"><input name="branch" type="text" class="Sytle_text" id="branch">
            <input name="Submit" type="button" class="Style_button" value="搜索" onClick="Mycheck();"></td>
          <td width="12%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="17%"><div align="right">
            <input name="Submit3" type="button" class="Style_button_del" value="优秀职员"
			 onclick="JScript:window.open('personnel_best.asp','','width=417,height=259')">
          </div></td>
        </tr>
      </table>
    </form></td>
  </tr>
</table>		</td>
      </tr>
</table>


<table width="556" border="0" align="center" cellpadding="-2" cellspacing="-2">
  <tr>
    <td>
	<%if rs_personnel.eof then%>
	<table align="center" cellpadding="0" cellspacing="0">
	  <tr><td><span class="STYLE16">无员工信息！</span></td>
	  </tr></table>
	<%else%>
	<table width="97%" height="42"  border="1" align="right" cellpadding="-2" cellspacing="-2" bordercolor="#FFFFFF" bordercolorlight="#FFCCCC" bordercolordark="#FFFFFF">
      <tr>
        <td width="19%"><div align="center" class="STYLE16">用户名</div></td>
        <td width="18%"><div align="center"><span class="STYLE16">姓名</span></div></td>
        <td width="10%"><div align="center" class="STYLE16">权限</div></td>
        <td width="14%"><div align="center" class="STYLE16">部门</div></td>
        <td width="16%"><div align="center"><span class="STYLE16">职务</span></div></td>
        <td width="11%"><div align="center" class="STYLE16">详细资料</div></td>
        <td width="6%"><div align="center" class="STYLE16">修改</div></td>
        <td width="6%"><div align="center" class="STYLE16">删除</div></td>
      </tr>
      　<%
	    '分页
		rs_personnel.pagesize=6
		page1=CLng(Request("page1"))
		if page1<1 then page1=1
		rs_personnel.absolutepage=page1
		for i=1 to rs_personnel.pagesize
		%>
      <tr>
          <td><div class="STYLE14">&nbsp;<%=(rs_personnel("UserName"))%></div></td>
          <td><div class="STYLE14">&nbsp;<%=rs_Personnel("Name")%></div></td>
          <td><div class="STYLE14">&nbsp;<%=(rs_personnel("purview"))%></div></td>
          <td><div class="STYLE14">&nbsp;<%=(rs_personnel("branch"))%></div></td>
          <td><div class="STYLE14">&nbsp;<%=(rs_personnel("job"))%></div></td>
          <td><div align="center">
		  <A HREF=# onClick="JScrip:window.open('personnel_detail.asp?ID=<%=rs_personnel("ID")
		   %>','','width=556,height=400')">
		   <img src="../images/detail.gif" width="16" height="17" border="0"></A></div></td>
          <td><div align="center">
		  <% if trim(rs_user("purview"))="系统" then%>
	          <A HREF=# onClick="JScrip:window.open('personnel_modify.asp?ID=<%=rs_personnel("ID")
			   %>','','width=460,height=400')">
			   <img src="../images/modify.gif" width="12" height="12" border="0"></a>
	      <% else%>
	          <img src="../images/modify.gif" width="12" height="12" border="0">
	      <% end if %>
          </div></td>
          <td><div align="center">
	  	  <% if trim(rs_user("purview"))="系统" then%>
              <A HREF=# onClick="JScrip:window.open('personnel_del.asp?ID=<%=rs_personnel("ID")
			   %>','','width=460,height=400')">
			   <img src="../images/del.gif" width="16" height="16" border="0"></a>            
          <% else%>
	          <img src="../images/del.gif" width="16" height="16" border="0">
	      <% end if %> 
          </div></td>
      </tr>
        <% rs_personnel.movenext
		if rs_personnel.eof then exit for 
		next %>
      </table></td>
  </tr>
</table>
<table width="556" border="0" align="center" cellpadding="-2" cellspacing="-2">
  <tr>
    <td><div align="right">
      　<span class="STYLE16">
      <% if page1<>1 then %>
      <a href=<%=path1%>?page1=1 class="l">第一页</a>
		<a href=<%=path1%>?page1=<%=(page1-1)%> class="l">上一页</a> 
		<%end if 
		if page1<>rs_personnel.pagecount then %>
   		<a href=<%=path1%>?page1=<%=(page1+1)%> class="l">下一页</a> 
		<a href=<%=path1%>?page1=<%=rs_personnel.pagecount%> class="l">最后一页</a> 
		<%end if %>
		&nbsp; </span></div></td>
  </tr>
</table>
<%end if %>
		
		&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
<%
rs_user.Close()
Set rs_user = Nothing
%>	
