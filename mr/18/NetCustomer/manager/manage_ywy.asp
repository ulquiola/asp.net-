<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<!--#include file="../../NetCustomer/F_Display_z.asp" --><!--包含显示区域信息的过程文件-->
<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM DB_WorkerInfo"
rs.open sql,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>网上客户管理系统--管理员登录！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image: url(../Images/bg.gif);
}
.style1 {color: #FFFFFF}
.style2 {color: #a2bcc5}
-->
</style></head>

<body>
<table width="595" height="242" border="0" cellpadding="-2" cellspacing="-2" background="../Images/manage.gif">
  <tr>
    <td valign="top"><table width="595" height="400" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="10" height="94">&nbsp;</td>
        <td width="577">&nbsp;</td>
        <td width="10"><span class="style2"></span></td>
      </tr>
      <tr>
        <td height="258">&nbsp;</td>
        <td valign="top">
          <div align="center">
            <table width="97%" height="31" cellpadding="-2"  cellspacing="-2">
              <tr>
                <td width="3%"><div align="right"><img src="../Images/add.gif" width="18" height="18">
				</div></td>
                <td width="79%"><a href="#" 
				onClick="JScript:window.open('areaselect.asp','','width=398,height=270,status=yes')">
				&nbsp;添加业务员</a></td>
				<td width="13%"><div align="right"><img src="../Images/back.GIF"
				 width="14" height="14"></div></td>
                <td width="5%"><a href="default.asp">返回</a></td>
              </tr>
            </table>
              <table width="100%" height="47" border="1" cellpadding="-2"  cellspacing="-2" 
			  bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr bgcolor="#809EA4">
                <td width="11%"><div align="center" class="style1">业务员姓名</div></td>
                <td width="9%"><div align="center" class="style1">性别</div></td>
                <td width="36%"><div align="center" class="style1">工作区域</div></td>
                <td width="20%"><div align="center" class="style1">电话</div></td>
                <td width="14%"><div align="center" class="style1">修改工作区域</div></td>
                <td width="5%"><div align="center" class="style1">修改</div></td>
                <td width="5%"><div align="center" class="style1">删除</div></td>
              </tr>
            <%'分页'
			rs.pagesize=6
			page=CLng(Request("page"))
			if page<1 then page=1
			rs.absolutepage=page
			for i=1 to rs.pagesize %>
              <tr>
                  <td align="center"><A href="#"
				  onClick="JScript:window.open('detail_ywyinfo.asp?ywynumber=<%=rs("ywynumber")
				   %>','','width=598,height=400')"><%=rs("Name")%></a></td>
                  <td align="center"><%=(rs("sex"))%></td>
                  <td align="center"><%qyNO=rs("workqy")
				      call Display(qyNO)  '调用过程显示业务员的工作区域%></td>
                  <td>&nbsp;<%=rs("tel")%></td>
                  <td align="center"><a href="#" 
				  onClick="JScript:window.open('areamodify.asp?ywynumber=<%=rs("ywynumber")
				   %>','','width=389,height=270')"><img src="../Images/modify.gif" width="20" height="18" border="0">
				   </a></td>
                  <td align="center"><A href="#"
				  onClick="JScript:window.open('modify_ywyinfo.asp?ywynumber=<%=rs("ywynumber")
				  %>','','width=598,height=400')"><img src="../Images/modify.gif" width="20"
				   height="18" border="0"></A></td>
                  <td align="center"><A HREF="#" 
				  onClick="JScript:window.open('del_ywyinfo.asp?ywynumber=<%=rs("ywynumber")
				  %>','','width=598,height=400')"><img src="../Images/del.gif" width="22"
				   height="22" border="0"></A></td>
			  </tr>
            <% rs.movenext
			if rs.eof then exit for 
			next %>
            </table>
            <table width="97%" height="31" cellpadding="-2"  cellspacing="-2">
              <tr>
               <td><div align="right">
            <% if page<>1 then %><a href=<%=path%>?page=1>第一页</a>
			<a href=<%=path%>?page=<%=(page-1)%>>上一页</a> 
			<%end if 
			if page<>rs.pagecount then %>
   			<a href=<%=path%>?page=<%=(page+1)%>>下一页</a> 
			<a href=<%=path%>?page=<%=rs.pagecount%>>最后一页</a> 
			<%end if %>
			</div></td>
           </tr>
         </table>
          </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>