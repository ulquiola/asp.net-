<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../Conn/conn.asp" -->
<%
If (Session("UserName")="") Then %>
<script language="javascript">
alert("您登录的网页已经过期，请重新登录！");
window.close();
opener.close();
window.open("../index.asp")
</script>
<%End If%>
<%
Set rs_user = Server.CreateObject("ADODB.Recordset")
sql_User= "SELECT ID, UserName, purview FROM Tab_User WHERE UserName = '"&session("UserName")&"'"
rs_user.open sql_User,conn,1,3
Set rs_bbs = Server.CreateObject("ADODB.Recordset")
sql_bbs= "SELECT ID, Subject, Person, dDate FROM Tab_Placard ORDER BY dDate DESC"
rs_bbs.open sql_bbs,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>公告栏</title>
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
.STYLE11 {font-size: 9pt}
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
    <td height="488" valign="top" background="../../Images/main.gif"><table width="88%" height="84%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td valign="top"><table width="557" height="119" border="0" cellpadding="-2" cellspacing="-2" background="../images/bbs/bbs_01.gif">
          <tr valign="top">
            <td width="800" height="86"><table width="223" height="83" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="32" valign="middle"><div align="right"><img src="../../Images/isexists.gif" width="16" height="16"></div></td>
                <td width="191">&nbsp;<img src="../../Images/g.gif" width="69" height="19"></td>
              </tr>
            </table></td>
          </tr>
          <tr valign="top">
            <td>&nbsp;</td>
          </tr>
        </table>
          <table width="672" border="0" cellpadding="-2" cellspacing="-2">
            <tr>
              <td width="817" valign="top"><table width="557" height="54"  border="0" 
	cellpadding="-2" cellspacing="-2" background="../images/bbs/bbs_02.gif">
                  <tr>
                    <td width="8%" height="33" align="center" valign="middle">&nbsp;</td>
                    <td width="19%" align="center">&nbsp;</td>
                    <td width="6%" height="33"><div align="center" class="style10"> <img src="../../images/add.gif" width="20" height="18"></div></td>
                    <td width="37%" height="33"><div align="left" class="STYLE11">
                        <% if trim(rs_user("purview"))<>"只读" then%>
                        <a href="#" class="STYLE11" onClick="JScript:window.open('bbs_add.asp','','width=560,height=397')"> 添加新公告</a>
                        <% else%>
                      添加新公告
                      <% end if%>
                    </div></td>
                    <td width="30%" height="33"><div align="center"></div></td>
                  </tr>
                  <tr>
                    <td height="21" colspan="6">&nbsp;</td>
                  </tr>
                </table>
				<%
				if rs_bbs.eof then
				response.Write("暂无公告信息!!")
				else
				%>
                  <table width="97%" height="42"  border="1" align="right" cellpadding="-2" cellspacing="-2" bordercolor="#FFFFFF" bordercolorlight="#FFCCCC" bordercolordark="#FFFFFF">
                    <tr>
                      <td width="54%"><div align="center" class="STYLE11">公告主题</div></td>
                      <td width="10%"><div align="center" class="STYLE11">公布人</div></td>
                      <td width="21%"><div align="center" class="STYLE11">公布时间</div></td>
                      <td width="8%"><div align="center" class="STYLE11">修改</div></td>
                      <td width="7%"><div align="center" class="STYLE11">删除</div></td>
                    </tr>
                    <%    
		    '分页
			rs_bbs.pagesize=7
			page1=CLng(Request("page1"))
			if page1<1 then page1=1
			rs_bbs.absolutepage=page1
			for i=1 to rs_bbs.pagesize %>
                    <tr>
                      <td><a href="#" onClick="JScript:window.open('bbs_detail.asp?ID=<%=rs_bbs("ID")%>','','width=550,height=397')">&nbsp;<%=rs_bbs("Subject")%></a></td>
                      <td><div align="center" class="STYLE11"><%=rs_bbs("Person")%></div></td>
                      <td class="STYLE11"><div align="center" class="style10"><%=rs_bbs("dDate")%></div></td>
                      <td class="STYLE11"><div align="center" class="STYLE11">
                          <% if trim(rs_user("purview"))<>"只读" then%>
                          <a href="#" onClick="JScript:window.open('bbs_modify.asp?ID=<%= rs_bbs("ID")%>','','width=550,height=397')"> <img src="../../images/modify.gif" width="12" height="12" border="0"></a>
                          <% else%>
                          <img src="../../images/modify.gif" width="12" height="12" border="0">
                          <% end if %>
                      </div></td>
                      <td><div align="center" class="STYLE11">
                          <% if trim(rs_user("purview"))<>"只读" then%>
                          <a href="#" onClick="JScript:window.open('bbs_del.asp?ID=<%=rs_bbs("ID")%>','','width=550,height=397')"><img src="../../images/del.gif" width="16" height="16" border="0"></a>
                          <% else%>
                          <img src="../../images/del.gif" width="16" height="16" border="0">
                          <% end if %>
                      </div></td>
                    </tr>
                    <% rs_bbs.movenext
			if rs_bbs.eof then exit for 
			next 
		%>
                </table>
	<%end if%>			
				
				</td>
            </tr>
          </table>
          <table width="674" border="0" cellspacing="-2" cellpadding="-2">
            <tr>
              <td><div align="right">
                  <% if page1<>1 then %>
                <a href=<%=path1%>?page1=1 class="STYLE11">第一页</a> <a href=<%=path1%>?page1=<%=(page1-1)%> class="STYLE11">上一页</a>
                  <%end if 
			if page1<>rs_bbs.pagecount then %>
                  <a href=<%=path1%>?page1=<%=(page1+1)%> class="STYLE11">下一页</a> <a href=<%=path1%>?page1=<%=rs_bbs.pagecount%> class="STYLE11">最后一页</a>
                  <%end if %>
                &nbsp; </div></td>
            </tr>
          </table></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>