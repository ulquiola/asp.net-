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
<title>员工信息！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
.style10 {color: #669999}
.style3 {color: #C60001}
.STYLE12 {color: #C60001; font-size: 9pt; }
-->
</style></head>

<body>
<table width="556" border="0" align="center" cellpadding="-2" cellspacing="-2">
  <tr>
    <td>
	<%if rs_personnel.eof then%>
	<table align="center"><tr><td>无员工信息！</td></tr></table>
	<%else%>
	<table width="97%" height="42"  border="1" align="right" cellpadding="-2" cellspacing="-2" bordercolor="#FFFFFF" bordercolorlight="#FFCCCC" bordercolordark="#FFFFFF">
      <tr>
        <td width="19%"><div align="center" class="STYLE12">用户名</div></td>
        <td width="18%"><div align="center"><span class="STYLE12">姓名</span></div></td>
        <td width="10%"><div align="center" class="STYLE12">权限</div></td>
        <td width="14%"><div align="center" class="STYLE12">部门</div></td>
        <td width="16%"><div align="center"><span class="STYLE12">职务</span></div></td>
        <td width="11%"><div align="center" class="STYLE12">详细资料</div></td>
        <td width="6%"><div align="center" class="STYLE12">修改</div></td>
        <td width="6%"><div align="center" class="STYLE12">删除</div></td>
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
          <td><div class="style10">&nbsp;<%=(rs_personnel("UserName"))%></div></td>
          <td><div class="style10">&nbsp;<%=rs_Personnel("Name")%></div></td>
          <td><div class="style10">&nbsp;<%=(rs_personnel("purview"))%></div></td>
          <td><div class="style10">&nbsp;<%=(rs_personnel("branch"))%></div></td>
          <td><div class="style10">&nbsp;<%=(rs_personnel("job"))%></div></td>
          <td><div align="center">
		  <A HREF=# onClick="JScrip:window.open('personnel_detail.asp?ID=<%=rs_personnel("ID")
		   %>','','width=556,height=400')">
		   <img src="../images/detail.gif" width="16" height="17" border="0"></A></div></td>
          <td><div align="center">
		  <% if trim(rs_user("purview"))="系统" then%>
	          <A HREF=# onClick="JScrip:window.open('personnel_modify.asp?ID=<%=rs_personnel("ID")
			   %>','','width=556,height=400')">
			   <img src="../images/modify.gif" width="12" height="12" border="0"></a>
	      <% else%>
	          <img src="../images/modify.gif" width="12" height="12" border="0">
	      <% end if %>
          </div></td>
          <td><div align="center">
	  	  <% if trim(rs_user("purview"))="系统" then%>
              <A HREF=# onClick="JScrip:window.open('personnel_del.asp?ID=<%=rs_personnel("ID")
			   %>','','width=556,height=400')">
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
      　<% if page1<>1 then %><a href=<%=path1%>?page=1 class="l">第一页</a>
		<a href=<%=path1%>?page1=<%=(page1-1)%> class="l">上一页</a> 
		<%end if 
		if page1<>rs_personnel.pagecount then %>
   		<a href=<%=path1%>?page1=<%=(page1+1)%> class="l">下一页</a> 
		<a href=<%=path1%>?page1=<%=rs_personnel.pagecount%> class="l">最后一页</a> 
		<%end if %>
		&nbsp; </div></td>
  </tr>
</table>
<%end if %>
</body>
</html>
