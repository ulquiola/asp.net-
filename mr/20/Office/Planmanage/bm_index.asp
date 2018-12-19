<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../conn/conn.asp"-->
<%
set rs=server.CreateObject("adodb.recordset")
sql="select * from Tab_bm order by time1 desc"
rs.open sql,conn,1,3
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>部门计划管理页面</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
.STYLE2 {
	font-size: 10pt;
	color: #000000;
}
a:link {
	text-decoration: none;
	color: #000000;
}
a:visited {
	text-decoration: none;
	color: #000000;
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
    <td height="488" valign="middle" background="../Images/main.gif"><table width="87%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="4%" valign="middle"><div align="right"><img src="../Images/isexists.gif" width="16" height="16" /></div></td>
        <td width="80%" valign="bottom"><div align="left"><span class="STYLE2">&nbsp;</span><img src="../Images/bm.gif" width="68" height="18" /></div></td>
        <td width="16%" valign="bottom" class="STYLE2">&nbsp;</td>
      </tr>
    </table>
      <br>
            <table width="91%" height="84%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="32"><table width="99%" height="32" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td background="../Images/waichubei.gif"><table width="99%" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="4%" valign="bottom">&nbsp;</td>
                <td width="72%" valign="bottom">&nbsp;</td>
                <td width="7%" valign="bottom"><div align="right"><img src="../Images/add.gif" width="20" height="18" /></div></td>
                <td width="17%" valign="bottom" class="STYLE2">&nbsp;<a href="#" onclick="Javascript:window.open('bm_add.asp','','width=560,height=360')">添加部门计划</a></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td valign="top">
		 <%
		If(rs.Eof)Then
			Response.Write("暂无部门计划信息!")
		Else
			
		%>
		<table width="728" height="55" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="23" valign="top"><div align="center" class="STYLE2">部门名称</div></td>
            <td><div align="center" class="STYLE2">计划主题</div></td>
            <td><div align="center" class="STYLE2">计划内容</div></td>
            <td><div align="center" class="STYLE2">发表日期</div></td>
            <td><div align="center" class="STYLE2">删除</div></td>
          </tr>
		  <%
		  '分页
		  rs.pagesize=8
		  page1=clng(request("page1"))
		  if page1<1 then page1=1
		  rs.absolutepage=page1
		  for i=1 to rs.pagesize
		  kk=rs("title")
		  k=rs("content")
		  %>
          <tr>
            <td height="23" class="STYLE2"><div align="center"><%=rs("name1")%>&nbsp;</div></td>
            <td height="23" class="STYLE2"><div align="center">
			<%if len(kk)>8 then%>
			<a href="#" onClick="Javascript:window.open('bm_xianxi.asp?ID=<%=rs("ID")%>','','width=456,height=459')"><%=replace(left(kk,8)&"....",chr(13),"<br>")%></a>
			<%else%>
			<a href="#" onClick="Javascript:window.open('bm_xianxi.asp?ID=<%=rs("ID")%>','','width=456,height=459')"><%=kk%></a>
			<%end if%>
		    &nbsp;
			</div></td>
            <td height="23" class="STYLE2"><div align="center">
			<%if len(k)>10 then%>
			<%=replace(left(k,10)&"....",chr(13),"<br>")%>
			<%else%>
			<%=k%>&nbsp;
			<%end if%>
			</div></td>
            <td height="23" class="STYLE2"><div align="center"><%=rs("time1")%>&nbsp;</div></td>
            <td height="23"><div align="center"><a href="bm_del.asp?id=<%=rs("id")%>" onclick="return confirm('是否确认删除?')"><img src="../Images/del.gif" width="16" height="16" border="0"></a></div></td>
          </tr>
		  <%
	  rs.movenext
	  if rs.eof then exit for 
	  next
	  %>
        </table>
		<table width="728" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td height="25" class="STYLE2"><div align="right" class="STYLE2">
                  <% if page1<>1 then%>
              <a href=<%=path1%>?page1=1 class="STYLE2">第一页</a> <a href=<%=path1%>?page1=<%=(page1-1)%> class="STYLE2"> 上一页</a>
              <%end if %>
              <%if page1<>rs.pagecount then%> 
              <a href=<%=path1%>?page1=<%=(page1+1)%> class="STYLE2">下一页</a> <a href=<%=path1%>?page1=<%=rs.pagecount%> class="STYLE2" >最后一页</a>
              <%end if%></div></td>
              </tr>
          </table>
		  <%end if%>
		</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
