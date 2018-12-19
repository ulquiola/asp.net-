<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<%
'查询所有管理员信息
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM DB_manager"
rs.open sql,conn,1,3
'查询当前登录的管理员信息
If (Session("UserName") <> "") Then 
  UserName = Session("UserName")
End If
Set rs_purview = Server.CreateObject("ADODB.Recordset")
sql_M = "SELECT * FROM DB_manager WHERE Name = '" & Replace(UserName, "'", "''") & "'"
rs_purview.open sql_M,conn,1,3
session("purview")=rs_purview("purview")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>网上客户管理系统--管理员登录！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image: url(Images/bg.gif);
}
.style1 {color: #FFFFFF}
.style2 {color: #a2bcc5}
-->
</style></head>

<body>
<table width="595" height="242" border="0" cellpadding="-2" cellspacing="-2"
 background="../Images/manage.gif">
  <tr>
    <td valign="top"><table width="595" height="400" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="15" height="94">&nbsp;</td>
        <td width="561">&nbsp;</td>
        <td width="17"><span class="style2"></span></td>
      </tr>
      <tr>
        <td height="258">&nbsp;</td>
        <td valign="top">
  
          <div align="center">
            <table width="97%" height="31" cellpadding="-2"  cellspacing="-2">
              <tr>
                <td width="4%"><div align="right"><img src="../Images/add.gif" width="18"
				 height="18"></div></td>
                <td width="48%">
                  <div align="left">
				  <% if rs_purview("purview")="系统" then %>
				  <a href="#" onClick="JScript:window.open('add_manager.asp','','width=398,height=270')">
				  &nbsp;添加管理员</a>
				  <%else%>
				  &nbsp;添加管理员
				  <% end if %>
				  </div></td>
                <td width="41%"><div align="right"><img src="../Images/back.GIF" width="14" height="14"></div></td>
                <td width="7%"><a href="default.asp">返回</a></td>
              </tr>
            </table>
              <table width="97%" height="47" border="1" cellpadding="-2"  cellspacing="-2"
			   bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr bgcolor="#809EA4">
                <td width="28%"><div align="center" class="style1">管理员名称</div></td>
                <td width="22%"><div align="center"><span class="style1">权限</span></div></td>
                <td width="17%"><div align="center" class="style1">修改密码</div></td>
                <td width="16%"><div align="center"><span class="style1">修改权限</span></div></td>
                <td width="17%"><div align="center" class="style1">删除</div></td>
              </tr>
            <%'分页显示管理员信息
			rs.pagesize=7
			page=CLng(Request("page"))
			if page<1 then page=1
			  rs.absolutepage=page
			for i=1 to rs.pagesize %>
              <tr>
               <td><span class="style1">&nbsp;</span><%=(rs.Fields.Item("Name").Value)%></td>
               <td><div align="center"><%=(rs.Fields.Item("purview").Value)%></div></td>
			   <% if rs_Purview("Name")=rs("name") or rs_purview("purview")="系统" then %>
                  <td><div align="center"><span class="style1"><A HREF="#"
				   onClick="JScript:window.open('modify_manager_PWD.asp?ID=<%= rs("ID")
				    %>','','width=398,height=270')"><img src="../Images/modify.gif"
					 width="20" height="18" border="0"></A></span></div></td>
               <%else%>
				  <td><div align="center">
				  <img src="../Images/modify.gif" width="20" height="18" border="0">
				  </A></span></div></td>
               <%end if%>
			   <% if rs_purview("purview")="系统" then %>
                  <td><div align="center"><span class="style1"><A HREF="#"
				   onClick="JScript:window.open('modify_manager_Pur.asp?ID=<%= rs("ID")
				    %>','','width=398,height=270')"><img src="../Images/modify.gif" width="20"
					 height="18" border="0"></A></span></div></td>
                  <td><div align="center"><span class="style1"><A HREF="#"
				   onClick="JScript:window.open('del_manager.asp?ID=<%= rs("ID")
				    %>','','width=389,height=270')"><img src="../Images/del.gif" width="22"
					 height="22" border="0"></A></span></div></td>
			  <% else %>
                  <td><div align="center"><span class="style1">
				  <img src="../Images/modify.gif" width="20" height="18" border="0">
				  </span></div></td>
                  <td><div align="center"><span class="style1">
				  <img src="../Images/del.gif" width="22" height="22" border="0">
				  </span></div></td>
              <% end if %>
			  </tr>
              <% rs.movenext
			  if rs.eof then exit for 
			next %>
            </table>
            <table width="97%" height="31" cellpadding="-2"  cellspacing="-2">
              <tr>
                <td><div align="right">
            <% if page<>1 then %><a href=<%=path%>?page=1 class="l">第一页</a>
			   <a href=<%=path%>?page=<%=(page-1)%> class="l">上一页</a> 
			<%end if 
			if page<>rs.pagecount then %>
   			   <a href=<%=path%>?page=<%=(page+1)%> class="l">下一页</a> 
			   <a href=<%=path%>?page=<%=rs.pagecount%> class="l">最后一页</a> 
			<%end if %>
			</div></td>
            </tr>
          </table>
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
