<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"--><!--包含数据库连接文件-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")			'创建记录集
sql="Select * From tb_User Where Status<>'管理员' order by id desc"		'查询数据
rs.Open sql,conn,1,3									'打开记录集
%>
<html>
<head>
<title>用户管理！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Css/style.css" rel="stylesheet">
<script src="JS/fun.js"></script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.STYLE1 {
	font-size: 10pt;
	font-weight: bold;
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
<body bgcolor="B9DFFF">
<table width="777" height="341"  border="0" align="center" cellpadding="-2" cellspacing="-2">
  <tr>
    <td valign="top"><table width="100%" border="0" cellpadding="-2" cellspacing="-2" class="tableBorder">
      <tr>
        <td></td>
        </tr>
      <tr>
        <td></td>
        </tr>
      <tr>
        <td height="2"></td>
      </tr>
      <tr>
        <td valign="top">
<%if rs.eof then
response.Write("暂无用户信息")
else
%>
      <table width="100%" height="205"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder">
        <tr>
          <td width="13" height="32" bgcolor="#4EAEE1">&nbsp;</td>
          <td width="689" bgcolor="#4EAEE1" style="color:#FFFFFF"> ≡<span class="STYLE1"><font color="#FFFFFF">当前位置：查看用户信息</font></span> ≡ </td>
          <td width="73" bgcolor="#4EAEE1">&nbsp;</td>
        </tr>
        <tr valign="top">
          <td height="171" colspan="3"><table width="100%" height="170"  border="0" cellpadding="-2" cellspacing="-2">
            <tr>
              <td height="10" valign="top"></td>
            </tr>
            <tr>
              <td valign="top"><table width="100%" height="74"  border="2" cellpadding="0" cellspacing="0" bordercolor="#4EAEE1" bordercolorlight="#FFFFFF" bgcolor="#FFFFFF" class="tableBorder_B">
                <tr align="center">
                  <td width="112">用户名</td>
                  <td width="102" height="25">真实姓名</td>
                  <td width="32">性别</td>
                  <td width="108">生日</td>
                  <td width="210">Email地址</td>
                  <td width="78">QQ</td>
                  <td width="56">发帖数量</td>
                  <td colspan="2">操作</td>
                </tr>
		<%
		'分页显示用户信息
		rs.pagesize=15							'设置分页显示时，每页显示记录的数目
		page1=CLng(Request("page1"))				'将获取到的记录数转换为整数
		if page1<1 then page1=1						'为page变量设置默认值
		rs.absolutepage=page1						'在进行分页显示时，指定指针所在的页
		for i=1 to rs.pagesize						'循环显示记录信息
		ID=rs("ID")									'获取记录集中ID的值并赋给ID变量
		%>
					<script language="javascript">
				  	function del(ID){				//创建del自定义函数
					if(confirm("真的要删除该用户吗？")){	//弹出提示对话框
						window.location.href="User_del.asp?UID="+ID;//跳转到指定的页面
					 }
					}
				  </script>
                <tr>
                  <td align="center"><%=rs("UserName")%></td>
                  <td height="25" align="center"><%=rs("TrueName")%></td>
                  <td align="center"><%=rs("Sex")%></td>
                  <td align="center"><%=rs("Birthday")%></td>
                  <td align="center"><%=rs("Email")%></td>
                  <td align="center"><%=rs("QQ")%>&nbsp;</td>
                  <td align="center"><%=rs("SendRatio")%></td>
                  <td width="37" align="center"><a href="#" onClick="del(<%=ID%>)"><img src="../images/del.gif" width="22" height="22" border="0" /></a></td>
                  <td width="20" align="center"><a href="modify_statc.asp?id=<%=ID%>"><img src="../Images/modify.gif" width="20" height="18" border="0"></a></td>
                </tr>
		<%
			  rs.MoveNext					'向下移动记录指针
			  If rs.Eof Then exit For 		'退出For循环
		      Next
		%> 
              </table>
                <table width="100%" height="26"  border="0" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
	<%If not(rs.Eof and rs.Bof) Then%>	
    <td width="34%" align="left">&nbsp;
	      		<% if page1<>1 then %><a href=?page1=1 class="white">第一页</a>
				<a href=?page1=<%=(page1-1)%> class="white">上一页</a> 
			<%
			end if 
			if page1<>rs.pagecount then %>
   				<a href=?page1=<%=(page1+1)%> class="white">下一页</a> 
				<a href=?page1=<%=rs.pagecount%> class="white">最后一页</a> 
		  <%end if%>
		  </td>
    <td width="66%" align="right" class="word_grey">[<%=page1%>/<%=rs.PageCount%>]&nbsp;&nbsp;每页<%=rs.PageSize%>条&nbsp;&nbsp;共<%=rs.RecordCount%>位用户&nbsp;</td>
		<%End If%>	
                </table></td>
            </tr>
            <tr>
              <td height="10" valign="top"></td>
            </tr>
          </table></td>
        </tr>
      </table>
<%end if%>
    </td>
  </tr>
</table>
</body>
</html>
<%
rs.close()				'关闭记录集
Set rs=Nothing			'释放Recordset对象占用的所有资源
%>