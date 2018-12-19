<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/conn.asp" -->
<!--#include file="../../NetCustomer/F_Display_z.asp" -->
<%
'查询业务员姓名
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM Q_Sell WHERE name = '"&session("username")&"'"
rs.open sql,conn,1,3
'查询业务员信息
Set rs_user = Server.CreateObject("ADODB.Recordset")
sql_user = "SELECT ywynumber,name,workqy FROM DB_WorkerInfo WHERE name='"&session("username")&"'"
rs_user.open sql_user,conn,1,3
'汇总销售金额
Set rs_money = Server.CreateObject("ADODB.Recordset")
sql_M = "SELECT sum(zmoney) as sMoney FROM Q_sell where name='"&session("username")&"'"
rs_money.open sql_M,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>销售信息查询！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image: url(../Images/bg_top.gif);
}
.style1 {color: #FFFFFF}
.style2 {color: #a2bcc5}
.style3 {color: #990000}
-->
</style></head>
<script language="javascript">
function mycheck(){
if (form1.sDate.value=="")
{alert("请输入查询起始日期！");form1.sDate.focus();return;}
if (form1.eDate.value=="")
{alert("请输入查询结束日期！");form1.eDate.focus();return;}
form1.submit();
}
</script>
<body><form method="post" action="quest_bottom.asp" target="info" name="form1">
<table width="595" height="100" border="0" cellpadding="-2" cellspacing="-2" background="../Images/manage_01_q.gif">
  <tr>
    <td valign="top">&nbsp;</td>
  </tr>
</table>
<table width="595" height="100" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
  <tr>
    <td valign="top"><table width="97%" height="24" cellpadding="-2"  cellspacing="-2">
      <tr>
        <td width="4%"><div align="right"><img src="../Images/add.gif" width="18" height="18"></div></td>
        <td width="71%"><div align="left">&nbsp;业务员：<%=session("UserName")%>　　工作区域：
                <%qyNO=rs_user("workqy")
				  call display(qyNO)%>
        </div></td>
        <td width="25%"><div align="right">总销售额为：<span class="style3">
		<%if Isnull(rs_money("sMoney")) then 
		sMoney=0 
		else 
		sMoney=rs_money("sMoney")
		end if%>
		<%=sMoney%>(元)</span></div></td>
      </tr>
    </table>
	
      <table width="595" height="27" cellpadding="-2" cellspacing="-2">
        <tr>
          <td width="43">
            <div align="right">
              <input name="F" type="checkbox" id="F" value="1" checked>
            </div></td>
          <td width="114"><div align="center">请选择查询条件：</div></td>
          <td width="83"><select name="field" id="field">
            <option value="khname" selected>客户名称</option>
            <option value="spname">商品名称</option>
          </select></td>
          <td width="61"><select name="TJ" id="TJ">
            <option value="=" selected>等于</option>
            <option value="like">Like</option>
          </select></td>
          <td width="147"><input name="INvalue" type="text" class="Sytle_auto" id="INvalue"></td>
          <td width="145"><input name="Button" type="button" class="Style_button" value="查询" onClick="mycheck()">
		  <input name="Button" type="button" class="Style_button" value="返回" onClick="JScript:window.parent.location.href='default.asp'"></td>
        </tr>
      </table>
      <table width="595" height="27" cellpadding="-2" cellspacing="-2">
        <tr>
          <td width="42"><div align="right">
            <input name="D" type="checkbox" id="D" value="1">
          </div></td>
          <td width="115"><div align="right">请输入起始日期：</div></td>
          <td width="119"><input name="sDate" type="text" class="Sytle_text" id="sDate" value="2004-5-1"></td>
          <td width="65">结束日期：</td>
          <td width="128"><input name="eDate" type="text" class="Sytle_text" id="eDate" value="<%=date()%>"></td>
          <td width="124">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</form></body>
</html>
