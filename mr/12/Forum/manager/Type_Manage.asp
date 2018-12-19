<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"--><!--包含数据库连接文件-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")			'创建记录集
SQL="Select * from view_Type"							'查询数据
rs.Open SQL,conn,1,3									'打开记录集
%>
<html>
<head>
<title>编辑讨论区</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Css/style.css" rel="stylesheet">
<script src="../JS/fun.js"></script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>
<body  bgcolor="B9DFFF">
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
      <table width="100%" height="205"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder">
        <tr>
          <td width="13" height="32" bgcolor="#4EAEE1">&nbsp;</td>
          <td width="689" bgcolor="#4EAEE1" style="color:#FFFFFF">&nbsp;&nbsp;<font color="#FFFFFF"><strong>当前位置：编辑讨论区</strong></font></td>
          <td width="73" bgcolor="#4EAEE1">&nbsp;</td>
        </tr>
        <tr valign="top">
          <td height="171" colspan="3"><table width="100%" height="189"  border="0" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
            
<%If rs.eof and rs.bof Then%>
            <tr>
			<td align="center">暂无讨论区信息！</td>
            </tr>
<%Else%>
            <tr>
              <td valign="top"><table width="100%"  border="1" cellpadding="-2" cellspacing="-2" bordercolor="#4EAEE1" bordercolorlight="#FFFFFF" class="tableBorder_B">
                <tr align="center">
                  <td width="326" height="20">讨论区名称</td>
                  <td width="204">版主</td>
                  <td width="301">添加时间</td>
                  <td width="82">修改</td>
                  <td width="85">删除</td>
                </tr>
		<%
		'分页
		rs.pagesize=15							'设置分页显示时，每页显示记录的数目
		page=CLng(Request("page"))				'将获取到的记录数转换为整数
		if page<1 then page=1					'为page变量设置默认值
		rs.absolutepage=page					'在进行分页显示时，指定指针所在的页
		for i=1 to rs.pagesize					'循环显示记录信息
		ID=rs("ID")								'获取记录集中ID的值并赋给ID变量
		%>
					<script language="javascript">
				  	function del(ID){									//创建del自定义函数
					if(confirm("真的要删除该类别吗？")){				//弹出提示对话框
						window.location.href="Type_del.asp?ID="+ID;		//跳转到指定的页面
					 }
					}
				  </script>
                <tr>
                  <td align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[ <%=rs("TName")%> ]</td>
                  <td align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("owner")%></td>
                  <td align="left"><div align="center"><%=rs("CreateTime")%>&nbsp;</div></td>
                  <td align="center"><a href="#" onClick="window.open('modify_Type.asp?TypeID=<%=rs("ID")%>','','width=240,height=139')"><img src="../Images/modify.gif" width="20" height="18" border="0"></a></td>
                  <td align="center"><input type="image" src="../Images/del.gif" width="22" height="18" border="0" onClick="del(<%=ID%>)" class="noborder"></td>
                </tr>
		<%
			  rs.MoveNext					'向下移动记录指针
			  If rs.Eof Then exit For 		'退出For循环
		  Next
		%>
              </table>
                <table width="100%"  border="0" cellspacing="-2" cellpadding="-2">
	<%If not(rs.Eof and rs.Bof) Then%>	
    <td width="34%" height="27" align="left">&nbsp;
	      		<% if page<>1 then %><a href=?page=1 class="white">第一页</a>
				<a href=?page=<%=(page-1)%> class="white">上一页</a> 
			<%end if 
			if page<>rs.pagecount then %>
   				<a href=?page=<%=(page+1)%> class="white">下一页</a> 
				<a href=?page=<%=rs.pagecount%> class="white">最后一页</a> 
		  <%end if%>	</td>
    <td width="66%" align="right" class="word_grey">[<%=page%>/<%=rs.PageCount%>]&nbsp;&nbsp;每页<%=rs.PageSize%>条&nbsp;&nbsp;共<%=rs.RecordCount%>个讨论区&nbsp;</td>
		<%End If   
	%>	
                </table></td>
            </tr>
            <tr>
              <td height="10" valign="top"></td>
            </tr>
          </table></td>
<%end If%>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
<%
rs.close()
Set rs=Nothing
%>
