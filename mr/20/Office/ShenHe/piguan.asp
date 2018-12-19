<!--#include file=../conn/conn.asp-->
<%
set rs_user=server.CreateObject("ADODB.Recordset")
sql="select * from Tab_User where username='"&Session("UserName")&"'"
rs_user.open sql,conn,1,3
%>
<% if trim(rs_user("purview"))<>"系统" then
%>
 <script language="javascript">
	alert("对不起您没有审核批示权限!!");
	window.location.href='../main.asp';
	</script>
<% end if%>
<%
set rs=server.CreateObject("adodb.recordset")
sql="select * from Tab_shehe  order by time1 desc"
rs.open sql,conn,1,3
%>
<%
if request.QueryString("id")<>"" then
yyy="update Tab_shehe set shen=1 where ID="&request.QueryString("id")
conn.execute(yyy)
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style>
<link href="table.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
a:link {
	color: #0000FF;
	text-decoration: none;
}
a:visited {
	color: #0000FF;
	text-decoration: none;
}
a:hover {
	color: #000000;
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
.style2 {color: #FFFFFF}
.style4 {color: #FF0000}
.style5 {color: #000000}
.style7 {
	color: #FFFFFF;
	font-size: 16px;
	font-family: "黑体";
}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
-->
</style>

</head>

<body>
<table width="814" height="505" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="488" valign="top" background="../Images/main.gif"><table width="735" height="313" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="53" valign="bottom"><table width="323" height="22" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="38"><div align="right"><img src="../Images/isexists.gif" width="16" height="16"></div></td>
            <td width="285" valign="bottom">&nbsp;<img src="../Images/s.gif" width="68" height="20"></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="260" valign="bottom"><table width="658" height="236" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td><table width="85%" height="40%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td valign="bottom"><table width="698" height="218" border="0" align="center" cellspacing="0">
                    <tr>
                      <td height="31" valign="middle" background="../Images/waichubei.gif">&nbsp;</td>
                    </tr>
                    <tr>
                      <td valign="top">
<%
		If(rs.Eof)Then
			Response.Write("暂无申请信息!")
		Else
			
		%>
					  <table width="698" border="1" align="center" cellspacing="1" bgcolor="#E6F5FF" bordercolordark="#000000" bordercolorlight="#E6F5FF">
                          <tr>
                            <td width="149" height="31"><div align="center" class="style5">申请标题</div></td>
                            <td width="128"><div align="center" class="style5">申请内容</div></td>
                            <td width="137"><div align="center" class="style5">日期</div></td>
                            <td colspan="3"><div align="center" class="style5">管理</div></td>
                          </tr>
						  		  <%
		  '分页
		  rs.pagesize=8
		  page1=clng(request("page1"))
		  if page1<1 then page1=1
		  rs.absolutepage=page1
		  for i=1 to rs.pagesize
		  k=rs("content")
		  %>

                          <tr >
                            <td height="26" bgcolor="#E6F5FF"><span class="style4">&nbsp;&nbsp;&nbsp;→</span>&nbsp;
							<a href="#" onClick="Javascript:window.open('xianxi.asp?ID=<%=rs("ID")%>','','width=456,height=459')"><%=rs("title")%></a>
							</td>
                            <td height="26" bgcolor="#E6F5FF" ><div align="center">			
							<%if len(k)>10 then%>
			<%=replace(left(k,10)&"....",chr(13),"<br>")%>
			<%else%>
			<%=k%>&nbsp;
			<%end if%>
</div></td>
                            <td bgcolor="#E6F5FF" ><div align="center"><%=rs("time1")%></div></td>
                            <td width="88" bgcolor="#E6F5FF" ><div align="center"><span class="style2"></span>
	   <%if rs("shen")=1 then%>
        已审核
        <% End If %>
        <%if rs("shen")=0 then%>
        <a href="piguan.asp?shen=1&id=<%=rs("id")%>" class="STYLE1" onClick="return confirm('确定审核吗？')">审核</a>
        <% End If %>
                             </div></td>
                            <td width="101" bgcolor="#E6F5FF" ><div align="center">
							<%if rs("shen")=1 then %> 
							修改
							<%end if%>
							<%if rs("shen")=0 then%> 
                            <a href="#" onClick="Javascript:window.open('update.asp?ID=<%=rs("ID")%>','','width=456,height=459')">
                             修改</a>
							 <%end if%>

							 </div>
							 </td>
                            <td width="76" bgcolor="#E6F5FF" ><div align="center"><a href="del.asp?id=<%=rs("id")%>" onClick="return confirm('是否确认删除?')">删除</a></div></td>
      </tr>
<%
	  rs.movenext
	  if rs.eof then exit for 
	  next
	  %>
                      </table>
         <table width="698" border="0" align="center" cellpadding="0" cellspacing="0">
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
            </table></td>
          </tr>
        </table>
          </td>
      </tr>
    </table>
    </td>
  </tr>
</table>
</body>
</html>
