<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file=../conn/conn.asp-->
<%
	if request("id")="" then
		id=session("id")
	else
		id=request("id")
	end if 
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from Tab_tongxun where id="&id
	rs.open sql,conn,1,1
	if request("del")<>"" then 
	call de
	end if
	function de
		sql4="delete from Tab_tongxunadd where id="&request("del")
		conn.execute(sql4)
	end function
	set rs1=server.CreateObject("adodb.recordset")
	sql1="select * from Tab_tongxunadd where name1="&rs("id")
	rs1.open sql1,conn,1,3
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<style type="text/css">
<!--
.style2 {color: #FFFFFF}
body,td,th {
	font-size: 12px;
}
a {
	font-size: 12px;
}
a:link {
	text-decoration: none;
	color: #0000FF;
}
a:visited {
	text-decoration: none;
	color: #0000FF;
}
a:hover {
	text-decoration: none;
	color: #0000FF;
}
a:active {
	text-decoration: none;
}
body {
	background-color: #ffffff;
	margin-left: 0px;
}
-->
</style>
</head>

<body>
<table width="519" height="225" border="1" align="center" cellspacing="0">
  <tr bgcolor="#FF9900">
    <td width="79" height="39" bgcolor="#6DBAFF"><div align="center" class="style2">通讯组名称</div></td>
    <td rowspan="2" valign="top" bgcolor="#6DBAFF">     <table width="100%" height="15%" border="0" cellspacing="1" bgcolor="#CCCCCC" >
      
      <tr bgcolor="#FF9900">
          <td width="18%" height="35" bgcolor="#6DBAFF" class="unnamed3"><div align="center"><span class="style2">名字</span></div></td>
          <td width="17%" bgcolor="#6DBAFF"><div align="center"><span class="style2">电话</span></div></td>
          <td width="19%" bgcolor="#6DBAFF"><div align="center">
            <div align="center" class="style2">QQ</div>
          </div></td>
          <td width="21%" bgcolor="#6DBAFF" style=" "><div align="center"><span class="style2">email</span></div></td>
          <td colspan="2" bgcolor="#6DBAFF" class="unnamed1"><div align="center"></div>          <div align="center"><span class="style2">管理</span></div></td>
        </tr>
        <%
	if not rs1.eof then
	max=rs1.recordcount
	for i=1 to max
	%>
        <tr>
          <td bgcolor="#FFFFFF" class="unnamed3"><div align="center">
		  <a href="#" onClick="JScript:window.open('lianxixian.asp?id=<%=rs1("id")%>','','width=350,height=500')"><%=rs1("name11")%></a></div></td>
          <td width="17%" bgcolor="#FFFFFF"><div align="center"><%=rs1("phone1")%></div></td>
          <td width="19%" bgcolor="#FFFFFF"><div align="center"><%=rs1("OICQ")%></div></td>
          <td width="21%" bgcolor="#FFFFFF" style=" "><div align="center"><a href="mailto:<%=rs1("email")%>"><%=rs1("email")%></a></div></td>
          <td width="12%" bgcolor="#FFFFFF" class="unnamed1"><div align="center">
 <a href="#" onClick="JScript:window.open('update.asp?id=<%=rs1("id")%>','','width=450,height=550')">修改</a>	  
		  </div></td>
          <td width="13%" bgcolor="#FFFFFF" class="unnamed2"><div align="center"><a href="del.asp?id=<%=rs1("id")%>" onClick="return confirm('是否确认删除?')">删除</a></div></td>
        </tr>
        <%
	 rs1.movenext
	 if rs1.eof then exit for
	 next
	 else 
	 response.Write("暂无信息")
	 end if
	 %>
          </table></td>
  </tr>
  <tr>
    <td height="173" valign="top" bgcolor="#6DBAFF"><table width="100%" border="0" bgcolor="#6DBAFF">
   
      <tr>
        <td height="30" valign="bottom"><div align="center"><%=rs("name1")%></div></td>
      </tr>
      <tr>
        <td height="2">
          <table width="80%" border="0" align="center" cellspacing="0">
            <tr>
              <td height="1" bgcolor="#FFFFFF"></td>
            </tr>
        </table></td>
      </tr>

    </table>
    </td>
  </tr>
</table>
</body>
</html>
<%session("id")=request("id")%>