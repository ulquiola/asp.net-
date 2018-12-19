<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../conn/conn.asp"-->
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
'查询收文箱内容
Set rs_lect = Server.CreateObject("ADODB.Recordset")
sql_lect="SELECT * FROM Tab_Send WHERE LPerson = '"&session("UserName")&"' ORDER BY sTime DESC"
rs_lect.open sql_lect,conn,1,3
'查询发文箱内容
Set rs_send = Server.CreateObject("ADODB.Recordset")
sql_send = "SELECT * FROM Tab_Send WHERE SPerson='"&session("UserName")&"' ORDER BY sTime DESC"
rs_send.open sql_send,conn,1,3
%>
<%
set rs_lperson = Server.CreateObject("ADODB.Recordset")
sql_lp="SELECT * FROM Tab_User WHERE  UserName='"&Session("UserName")&"'"
rs_lperson.open sql_lP,conn,1,3
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>收/发文管理页面</title>
<link href="../CSS/style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
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
<table width="814" height="520" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" background="../Images/main.gif">
	<table width="100%" height="74%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td>
		
		
		<table width="714" height="89" border="0" align="center" cellpadding="-2" cellspacing="-2">
  <tr valign="top">
    <td width="41" valign="middle"><div align="right"><img src="../Images/isexists.gif" width="16" height="16" /></div></td>
    <td width="673" valign="middle">&nbsp;<img src="../Images/sfw.gif" width="93" height="18" /></td>
  </tr>
</table>
<table width="557" border="0" align="center" cellpadding="-2" cellspacing="-2">
  <tr>
    <td width="817" valign="top"><table width="557" height="54"  border="0" cellpadding="-2" cellspacing="-2">
        <tr>
          <td width="8%" height="33" align="center" valign="middle">&nbsp;</td>
          <td width="19%" align="center">&nbsp;</td>
          <td width="6%" height="33"><div align="center" class="style10"><img src="../images/add.gif" width="20" height="18"></div></td>
          <td width="37%" height="33">
            <div align="left" class="style10">
			<%if rs_lperson("purview")="只读" then%>
			发送公文
			<%end if%>
			<%if rs_lperson("purview")<>"只读" then%>
			<a href="#" onClick="Javascript:window.open('send_add.asp','','width=460,height=397')">发送公文</a>
			<%end if%>
		  </div></td>
          <td width="30%" height="33"><div align="center"></div></td>
        </tr>
        <tr>
          <td height="21" colspan="6" class="style10">&nbsp;&nbsp;&nbsp;收文箱：</td>
        </tr>
    </table>
	<%if rs_lect.eof then%>
	<table align="center"><tr><td>收文箱为空！</td></tr></table>
	<%else%>
	<table width="97%" height="42"  border="1" align="right" cellpadding="-2" cellspacing="-2"
	 bordercolor="#FFFFFF" bordercolorlight="#FFCCCC" bordercolordark="#FFFFFF">
        <tr>
          <td width="8%"><div align="center" class="style3">状态</div></td>
          <td width="48%"><div align="center" class="style3">收文主题</div></td>
          <td width="16%"><div align="center" class="style3">发文人</div></td>
          <td width="22%"><div align="center"><span class="style3">发文时间</span></div></td>
          <td width="6%"><div align="center" class="style3">删除</div></td>
        </tr>
        <%'分页'
			rs_lect.pagesize=3
			page2=CLng(Request("page2"))
			if page2<1 then page2=1
				rs_lect.absolutepage=page2
			for i=1 to rs_lect.pagesize %>
        <tr>
          <td><div align="center">
		  <!--显示收文状态-->
		  <% if rs_lect("flag")="否" then %>
            	<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
				 codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.
				 cab#version=6,0,29,0" width="34" height="18">
              	<param name="movie" value="../swf/new.swf">
              	<param name="quality" value="high">
              	<param name="mwode" value="transparent">
              	<embed src="../swf/new.swf" width="34" height="18" quality="high" 
				pluginspage="http://www.macromedia.com/go/getflashplayer"
				 type="application/x-shockwave-flash" mwode="transparent"></embed>
				 </object>
			<% else %>
				<img src="../images/send/read.gif">
			<% end if %>
          </div></td>
          <td>
		  <div align="center" class="style10"><a href="#" onClick="JScript:window.open('owen_detail.asp?ID=<%=rs_lect("ID")
		  %>','','width=560,height=397')"><%=rs_lect("subject")%></a></div>
		  </td>
          <td><div align="center" class="style10"><%=rs_lect("SPerson")%></div></td>
          <td><div align="center" class="style10"><%=rs_lect("sTime")%></div></td>
          <td><div align="center"><a href="#" onClick="JScript:window.open('send_del.asp?ID=<%=rs_lect("ID")%>','','width=560,height=397')"><img src="../images/del.gif" width="16" height="16" border="0"></a>
		  </div></td>
        </tr>
        <% rs_lect.movenext
			if rs_lect.eof then exit for 
			next %>
        </table></td>
  </tr>
</table>
<table width="556" border="0" align="center" cellpadding="-2" cellspacing="-2">
  <tr>
    <td><div align="right">
      <% if page2<>1 then %><a href=<%=path2%>?page2=1>第一页</a>
			<a href=<%=path2%>?page2=<%=(page2-1)%>>上一页</a> 
	  <%end if 
	  if page2<>rs_lect.pagecount then %>
   			<a href=<%=path2%>?page2=<%=(page2+1)%>>下一页</a> 
			<a href=<%=path2%>?page2=<%=rs_lect.pagecount%>>最后一页</a> 
	  <%end if %>
	&nbsp; </div></td>
  </tr>
</table>
<% end if %>
<table width="557" height="69" border="0" align="center" cellpadding="-2" cellspacing="-2">
  <tr>
    <td valign="top"><table width="557" border="0" cellspacing="-2" cellpadding="-2">
      <tr>
        <td>&nbsp;&nbsp;&nbsp;<span class="style10">发文箱：</span></td>
      </tr>
    </table>
	<%if rs_send.eof then%>
	<table align="center"><tr><td>发文箱为空！</td></tr></table>
	<%else%>
      <table width="97%" height="42"  border="1" align="right" cellpadding="-2" cellspacing="-2" bordercolor="#FFFFFF" bordercolorlight="#FFCCCC" bordercolordark="#FFFFFF">
        <tr>
          <td width="8%"><div align="center" class="style3">状态</div></td>
          <td width="48%"><div align="center" class="style3">发文主题</div></td>
          <td width="16%"><div align="center" class="style3">收文人</div></td>
          <td width="22%"><div align="center"><span class="style3">发文时间</span></div></td>
        </tr>
        <%
		    '分页
			rs_send.pagesize=5
			page1=CLng(Request("page1"))
			if page1<1 then page1=1
			rs_send.absolutepage=page1
			for i=1 to rs_send.pagesize 
		%>
        <tr>
          <td><div align="center">
              <% if rs_send("flag")="否" then %>
              <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="34" height="18">
                <param name="movie" value="../swf/new.swf">
                <param name="quality" value="high">
                <param name="mwode" value="transparent">
                <embed src="../swf/new.swf" width="34" height="18" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" mwode="transparent"></embed>
              </object>
              <%else%>
              <img src="../images/send/read.gif">
            <%end if%>
            </div></td>
          <td>
		  <div align="center" class="style10">
		  <a href="#" onClick="JScript:window.open('send_detail.asp?ID=<%=rs_send("ID")%>','','width=560,height=397')"><%=rs_send("subject")%></a>
		  </div>
		  </td>
          <td><div align="center" class="style10"><%=rs_send("LPerson")%></div></td>
          <td><div align="center" class="style10"><%=rs_send("sTime")%></div></td>
        </tr>
        <%
		    rs_send.movenext
			if rs_send.eof then exit for 
			next 
		%>
            </table></td>
  </tr>
</table>
<table width="557" border="0" align="center" cellpadding="-2" cellspacing="-2">
  <tr>
    <td><div align="right">
      <% if page1<>1 then %><a href=<%=path1%>?page1=1>第一页</a>
			<a href=<%=path1%>?page1=<%=(page1-1)%>>上一页</a> 
			<%end if 
			if page1<>rs_send.pagecount then %>
   			<a href=<%=path1%>?page1=<%=(page1+1)%>>下一页</a> 
			<a href=<%=path1%>?page1=<%=rs_send.pagecount%>>最后一页</a> 
			<%end if %>
&nbsp; </div></td>
  </tr>
</table>
<%end if%>
		</td>
      </tr>
    </table></td>
  </tr>
</table></body>
</html>
