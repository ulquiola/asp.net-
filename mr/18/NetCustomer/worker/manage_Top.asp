<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/conn.asp" -->
<!--#include file="../../NetCustomer/F_Display_z.asp" -->
<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM Q_Sell WHERE name = '"&Session("UserName")&"'"
rs.open sql,conn,1,3
Set rs_user = Server.CreateObject("ADODB.Recordset")
sql_user= "SELECT ywynumber, name, workqy FROM DB_WorkerInfo WHERE name = '"&Session("UserName")&"'"
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
<body><form method="post" action="manage_bottom.asp" target="info" name="form1">
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
            <div align="right">            </div></td>
          <td width="438"><div align="left">
            <table width="433">
              <tr>
                <td colspan="2"><label>
                  <input name="TJ" type="radio" value="1" checked>
                  按客户名称</label></td>
                <td colspan="2"><input name="TJ" type="radio" value="2">
                  按商品名称</td>
              </tr>
              <tr>
                <td width="94"><label>
				<input type="radio" name="TJ" value="3">按销售日期</label></td>
                <td colspan="2">起始日期：
                  <input name="sDate" type="text" class="Sytle_text" id="sDate" value="2004-5-1"></td>
                <td width="170">结束日期：
                  <input name="eDate" type="text" class="Sytle_text" id="eDate2" value="<%=date()%>"></td>
              </tr>
            </table>
          </div></td>
          <td width="112"><input name="Butt_quest" type="submit" class="Style_button" id="Butt_quest" value="汇总">              
            <input name="Button" type="button" class="Style_button" value="返回" onClick="JScript:window.parent.location.href='default.asp'"></td></tr>
      </table>
      </td>
  </tr>
</table>
</form></body>
</html>
