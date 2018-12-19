<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../conn/conn.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
sql="select m.id,m.name1,isnull(d.c,0) as c from Tab_tongxun m left join (select name1,count(*) as c from Tab_tongxunadd group by name1) d on m.id=d.name1"
rs.open sql,conn,1,3
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>显示通讯组</title>
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
	color: #000000;
	text-decoration: none;
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
.style3 {color: #FF0000}
-->
</style></head>

<body>
<table width="814" height="505" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="488" valign="top" background="../Images/main.gif"><br>
        <table width="89%" height="75%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="55" valign="top"><table width="433" height="39" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="35" height="12" valign="bottom"></td>
            <td width="398" height="8" valign="bottom"></td>
          </tr>
          <tr>
            <td width="35" height="16" valign="middle"><div align="right"><img src="../Images/isexists.gif" width="16" height="16" /></div></td>
            <td height="34" valign="middle">&nbsp;<img src="../Images/tx.gif" width="84" height="19" /></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="36" valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="5%" background="../Images/waichubei.gif">&nbsp;</td>
            <td width="66%" background="../Images/waichubei.gif">&nbsp;</td>
            <td width="10%" background="../Images/waichubei.gif">&nbsp;</td>
            <td width="4%" background="../Images/waichubei.gif"><img src="../Images/add.gif" width="20" height="18" /></td>
            <td width="15%" background="../Images/waichubei.gif" class="STYLE2"><a href="#" onclick="JScript:window.open('tongxun_add.asp','','width=300,height=230')">添加通讯组类别</a></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td valign="top">
		<%
		If(rs.Eof)Then
			Response.Write("暂无通讯组信息")
		Else
		%>
          <table width="709" height="56" border="1" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="26"><div align="center" class="STYLE2">通讯组名称</div></td>
              <td width="212"><div align="center" class="STYLE2">通讯组内数量</div></td>
              <td colspan="2"><div align="center" class="STYLE2">管理通讯组</div></td>
            </tr>
            <%
		  '分页
		  rs.pagesize=8
		  page1=clng(request("page1"))
		  if page1<1 then page1=1
		  rs.absolutepage=page1
		  for i=1 to rs.pagesize
		  %>
            <tr>
              <td><div class="STYLE2"><span class="style3"> 　　　　→</span>&nbsp;<a href="#" onclick="JScript:window.open('tongxun_xianshi.asp?id=<%=rs("id")%>','','width=542,height=250')"><%=rs("name1")%></a></div></td>
              <td><div align="center" class="STYLE2">
			  &nbsp;<%=rs("c")%></div></td>
              <td width="116"><div align="center" class="STYLE2"><a href="#" onclick="JScript:window.open('tongxun_modify.asp?id=<%=rs("id")%>','','width=300,height=230')">修改</a></div></td>
              <td width="119"><div align="center" class="STYLE2"><a href="tongxun_del.asp?id=<%=rs("id")%>" onclick="return confirm('是否确认删除?')">删除</a></div></td>
            </tr>
            <%
	  rs.movenext
	  if rs.eof then exit for 
	  next
	  %>
          </table>
          <table width="709" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="25" class="STYLE2"><div align="right" class="STYLE2">
                  <% if page1<>1 then%>
                  <a href="<%=path1%>?page1=1" class="STYLE2">第一页</a> <a href="<%=path1%>?page1=<%=(page1-1)%>" class="STYLE2"> 上一页</a>
                  <%end if %>
                  <%if page1<>rs.pagecount then%>
                  <a href="<%=path1%>?page1=<%=(page1+1)%>" class="STYLE2">下一页</a> <a href="<%=path1%>?page1=<%=rs.pagecount%>" class="STYLE2" >最后一页</a>
                  <%end if%>
              </div></td>
            </tr>
          </table>
          <%end if%></td>
      </tr>
    </table></td>
  </tr>
</table>

</body>
</html>
