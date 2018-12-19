<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Conn/conn.asp" -->
<%
Set rs_user = Server.CreateObject("ADODB.Recordset")
sql_user= "SELECT ID, UserName, purview FROM Tab_User WHERE UserName = '"&session("UserName")&"'"
rs_user.open sql_user,conn,1,3

Set rs_personnel = Server.CreateObject("ADODB.Recordset")
sql_P= "SELECT ID,UserName,Name,purview,branch,job FROM Tab_User WHERE branch='"&_
Request("branch")&"' ORDER BY UserName ASC"
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
.STYLE11 {color: #C60001; font-size: 9pt; }
-->
</style></head>
<script language="javascript">
function Mycheck()
{
window.location.href='personnel_top.asp';
}
</script>
<body>
<table width="814" height="505" border="0" cellpadding="0" cellspacing="0" background="../Images/main.gif">
  <tr>
    <td height="488" valign="top" ><table width="87%" height="84%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="50"><table width="270" height="20" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="29"><div align="center"><img src="../Images/isexists.gif" width="16" height="16"></div></td>
            <td width="241"><img src="../Images/y.gif" width="100" height="19"></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td>
		<%if rs_personnel.eof then%>
          <table align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td><span class="STYLE16">暂无员工信息！</span></td>
            </tr>
          </table>
          <%else%>
          <table width="556" border="0" cellspacing="-2" cellpadding="-2">
            <tr>
              <td><table width="97%" height="42"  border="1" align="right" cellpadding="-2" cellspacing="-2" bordercolor="#FFFFFF" bordercolorlight="#FFCCCC" bordercolordark="#FFFFFF">
                  <tr>
                    <td width="19%"><div align="center" class="STYLE11">用户名</div></td>
                    <td width="18%"><div align="center"><span class="STYLE11">姓名</span></div></td>
                    <td width="10%"><div align="center" class="STYLE11">权限</div></td>
                    <td width="14%"><div align="center" class="STYLE11">部门</div></td>
                    <td width="16%"><div align="center"><span class="STYLE11">职务</span></div></td>
                    <td width="11%"><div align="center" class="STYLE11">详细资料</div></td>
                    <td width="6%"><div align="center" class="STYLE11">修改</div></td>
                    <td width="6%"><div align="center" class="STYLE11">删除</div></td>
                  </tr>
      <%
	     '分页
		rs_personnel.pagesize=2
		page1=CLng(Request.QueryString("page1"))
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
                  <td><div align="center"><a href=# 
		  onClick="JScrip:window.open('personnel_detail.asp?ID=<%=rs_personnel("ID")
		   %>','','width=556,height=400')"> <img src="../images/detail.gif" width="16" height="17" border="0"></a></div></td>
                  <td><div align="center">
                    <% if trim(rs_user("purview"))="系统" then%>
                    <a href=# onClick="JScrip:window.open('personnel_modify.asp?ID=<%=rs_personnel("ID")
			   %>','','width=456,height=400')"> <img src="../images/modify.gif" width="12" height="12" border="0"></a>
                    <% else%>
                    <img src="../images/modify.gif" width="12" height="12" border="0">
                    <% end if %>
                  </div></td>
                  <td><div align="center">
                    <% if trim(rs_user("purview"))="系统" then%>
                    <a href=# onClick="JScrip:window.open('personnel_del.asp?ID=<%=rs_personnel("ID")
			   %>','','width=456,height=400')"> <img src="../images/del.gif" width="16" height="16" border="0"></a>
                    <% else%>
                    <img src="../images/del.gif" width="16" height="16" border="0">
                    <% end if %>
                  </div></td>
                </tr>
                <%
		rs_personnel.movenext
		if rs_personnel.eof then exit for 
		next 
		%>
              </table></td>
            </tr>
          </table>
          <%end if %>
          <table width="556" border="0" cellspacing="-2" cellpadding="-2">
            <tr>
              <td><div align="right" class="STYLE11"> 　
  <% if page1<>1 then %>
   <a href=<%=path1%>?page1=1&branch=<%=Request("branch")%>>第一页</a> <a href=<%=path1%>?page1=<%=(page1-1)%>&branch=<%=Request("branch")%>>上一页</a>
   <%end if 
	if page1<>rs_personnel.pagecount then %>
   <a href=<%=path1%>?page1=<%=(page1+1)%>&branch=<%=Request("branch")%>>下一页</a> <a href=<%=path%>?page1=<%=rs_personnel.pagecount%>&branch=<%=Request("branch")%>>最后一页</a>
   <%end if %>

             </div></td>
            </tr>
            <tr>

              <td height="37"><div align="right">
                  <label>
                  <input type="submit" name="button" value="返回" onClick="Mycheck();">
                  </label>
              </div></td>
            </tr>
          </table>
		</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>