<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../Conn/conn.asp" -->
<%If Session("UserName")="" Then %>
<script language="javascript">
alert("您登录的网页已经过期，请重新登录！");
window.close();
opener.close();
window.open("login.asp")
</script>
<%End If%>
<%
Set rs_user = Server.CreateObject("ADODB.Recordset")
sql_user= "SELECT ID, UserName, purview FROM Tab_User WHERE UserName ='"&_
Session("UserName")&"'"
rs_user.open sql_user,conn,1,3
'查询会议信息
Set rs = Server.CreateObject("ADODB.Recordset")
sql= "SELECT ID, mTime, ZPerson, subject FROM Tab_meeting ORDER BY mTime DESC"
rs.open sql,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>会议管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
.style3 {color: #C60001}
.style10 {color: #669999}
.STYLE11 {
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
-->
</style></head>
<body>
<table width="814" height="505" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="787" height="488" valign="top" background="../../Images/main.gif"><table width="695" height="84%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td valign="top"><table width="557" height="116" border="0" cellpadding="-2" cellspacing="-2" background="../images/meeting/meeting_01.gif">
          <tr valign="top">
            <td width="800" height="52" valign="top"><table width="223" height="80" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="32"><div align="right"><img src="../../Images/isexists.gif" width="16" height="16"></div></td>
                <td width="191">&nbsp;<img src="../../Images/h.gif" width="69" height="18"></td>
              </tr>
            </table></td>
          </tr>
          <tr valign="top">
            <td>&nbsp;</td>
          </tr>
        </table>
          <table width="557" border="0" cellpadding="-2" cellspacing="-2">
            <tr>
              <td width="817" valign="top"><table width="557" height="54"  border="0" cellpadding="-2" cellspacing="-2" background="../images/meeting/meeting_02.gif">
                  <tr>
                    <td width="7%" height="33" align="center" valign="middle">&nbsp;</td>
                    <td width="23%" align="center">&nbsp;</td>
                    <td width="5%" height="33"><div align="center" class="style10"><img src="../../images/add.gif" width="20" height="18"></div></td>
                    <td width="47%" height="33"><div align="left" class="STYLE11">
                        <% if trim(rs_user("purview"))<>"只读" then%>
                        <a href="#" class="STYLE11" onClick="Javascript:window.open('meeting_add.asp','','width=560,height=397')">录入会议信息</a>
                        <% else%>
                      录入会议信息
                      <% end if%>
                    </div></td>
                    <td width="18%" height="33"><div align="center"></div></td>
                  </tr>
                  <tr>
                    <td height="21" colspan="6">&nbsp;</td>
                  </tr>
                </table>
                  <%if rs.eof then%>
                  <table align="center">
                    <tr>
                      <td class="STYLE11">暂无会议信息！</td>
                    </tr>
                  </table>
                <%else%>
                  <table width="95%" height="27"  border="1" align="right" cellpadding="-2" cellspacing="-2" bordercolor="#FFFFFF" bordercolorlight="#FFCCCC" bordercolordark="#FFFFFF">
                    <tr>
                      <td width="55%"><div align="center" class="STYLE11">会议主题</div></td>
                      <td width="13%"><div align="center" class="STYLE11">主持人</div></td>
                      <td width="25%"><div align="center" class="STYLE11">会议时间</div>
                          <div align="center" class="style3"></div></td>
                      <td width="7%"><div align="center" class="STYLE11">删除</div></td>
                    </tr>
                    <%'分页'
		rs.pagesize=9
		page1=CLng(Request("page1"))
		if page1<1 then page1=1
		rs.absolutepage=page1
		for i=1 to rs.pagesize 
		%>
                    <tr>
                      <td><a href="#" class="STYLE11" onClick="JScript:window.open('meeting_detail.asp?ID=<%=rs("ID") %>','','width=560,height=397')"><%=rs("subject")%></a></td>
                      <td><div align="center" class="STYLE11"><%=(rs.Fields.Item("ZPerson").Value)%></div></td>
                      <td><div align="center" class="STYLE11"><%=(rs.Fields.Item("mTime").Value)%> </div></td>
                      <td><div align="center">
                          <% if trim(rs_user("purview"))<>"只读" then%>
                          <A href="#" onClick="JScript:window.open('meeting_del_OK.asp?ID=<%=rs("ID")%>','','width=560,height=400')"><img src="../../images/del.gif" width="16" height="16" border="0"></a>
                          <% else%>
                          <img src="../../images/del.gif" width="16" height="16" border="0">
                          <% end if %>
                      </div></td>
                    </tr>
                    <% rs.movenext
		if rs.eof then exit for 
		next %>
                </table></td>
            </tr>
          </table>
          <table width="556" border="0" cellspacing="-2" cellpadding="-2">
            <tr>
              <td><div align="right">
                  <% if page1<>1 then %>
               <a href=<%=path1%>?page1=1 class="STYLE11">第一页</a> <a href=<%=path1%>?page1=<%=(page1-1)%> class="STYLE11">上一页</a>
                  <%
				  end if 
		          if page1<>rs.pagecount then 
		          %>
                  <a href=<%=path1%>?page1=<%=(page1+1)%> class="STYLE11">下一页</a> <a href=<%=path1%>?page1=<%=rs.pagecount%> class="STYLE11">最后一页</a>
                  <%end if %>
                &nbsp;</div></td>
            </tr>
          </table>
          <%end if%></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>