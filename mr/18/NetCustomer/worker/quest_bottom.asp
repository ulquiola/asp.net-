<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/conn.asp" -->
<!--#include file="../../NetCustomer/F_Display_z.asp" -->
<% '根据用户设置条件进行查询
Set rs = Server.CreateObject("ADODB.Recordset")
sql1=""
if request.Form("F")=1 and request.Form("D")=1 then
	if request.Form("TJ")="=" then
		sql2=" and "&request.Form("field") &" = '"& request.Form("INvalue") &"'" 
	else
		sql2=" and "&request.Form("field") &" like '%"& request.Form("INvalue") &"%'" 
	end if
	sql1=" and D_date between #"&request.Form("sDate")&"# and #"&request.Form("eDate")&"#"&sql2
else
	if request.Form("F")=1 then
		if request.Form("INvalue")<>"" then
			if request.Form("TJ")="=" then
				sql1=" and "&request.Form("field") &" = '"& request.Form("INvalue") &"'" 
			else
				sql1=" and "&request.Form("field") &" like '%"& request.Form("INvalue") &"%'" 
			end if
		end if
	end if
	if request.Form("D")=1 then
		sql1=" and D_date between #"&request.Form("sDate")&"# and #"&request.Form("eDate")&"#" 
	end if
end if
sql="SELECT * FROM Q_Sell WHERE name = '"&Session("UserName")&"'"&sql1
rs.open sql,conn,1,3
'查询业务员信息
Set rs_user = Server.CreateObject("ADODB.Recordset")
sql_user="SELECT ywynumber, name, workqy FROM DB_WorkerInfo WHERE name ='"&Session("UserName")&"'"
rs_user.open sql_user,conn,1,3
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
	background-image: url(../Images/bg.gif);
}
.style1 {color: #FFFFFF}
.style2 {color: #a2bcc5}
.style3 {color: #990000}
-->
</style></head>

<body>
<table width="595" height="142" border="0" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
  <tr>
    <td valign="top"><table width="595" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="10" height="40">&nbsp;</td>
        <td width="577" valign="top">          <div align="center">
		<%if not rs.eof then%>
            <table width="100%" border="1" cellpadding="-2"  cellspacing="-2" bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr bgcolor="#809EA4">
                <td width="23%"><div align="center" class="style1">客户名称</div></td>
                <td width="13%"><div align="center" class="style1">商品名称</div></td>
                <td width="11%"><div align="center" class="style1">包装</div></td>
                <td width="13%"><div align="center" class="style1"></div>                  
                <div align="center" class="style1">销售数量</div>                
                <div align="center" class="style1 style1">
                    <div align="center"></div>
                </div>                <div align="center" class="style1"></div></td>
                <td width="11%"><div align="center"><span class="style1">单价(元)</span></div></td>
                <td width="16%"><div align="center" class="style1">金额(元)</div></td>
                <td width="13%"><div align="center" class="style1">销售日期</div></td>
              </tr>
            <%'分页'
			rs.pagesize=5
			page=CLng(Request("page"))
			if page<1 then page=1
			rs.absolutepage=page
			for i=1 to rs.pagesize %>
              <tr>
                  <td><span class="style1">&nbsp;</span><%=(rs.Fields.Item("khname").Value)%></td>
                  <td><div align="center"><%=(rs.Fields.Item("spname").Value)%></div></td>
                  <td><div align="center"><%=(rs.Fields.Item("pack").Value)%></div></td>
                  <td><%=(rs.Fields.Item("sellsum").Value)%></td>
                  <td><div align="center"><%=(rs.Fields.Item("price").Value)%></div></td>
                  <td><div align="center"><%=(rs.Fields.Item("zmoney").Value)%></div></td>
                  <td><div align="center"><%=(rs.Fields.Item("D_date").Value)%></div></td>
			  </tr>
            <% rs.movenext
			if rs.eof then exit for 
			next %>
              </table>
			  <% end if %>
            <table width="97%" height="31" cellpadding="-2"  cellspacing="-2">
              <tr>
                <td><div align="right">
            <% if page<>1 then %><a href=<%=path%>?page=1 class="l">第一页</a>
			<a href=<%=path%>?page=<%=(page-1)%> class="l">上一页</a> 
			<%end if 
			if page<>rs.pagecount then %>
   			<a href=<%=path%>?page=<%=(page+1)%> class="l">下一页</a> 
			<a href=<%=path%>?page=<%=rs.pagecount%> class="l">最后一页</a> 
			<%end if %>
</div></td>
              </tr>
            </table>
			<% if rs.eof and rs.bof then%>
            <table width="595" height="45" cellpadding="-2" cellspacing="-2">
              <tr>
                <td><div align="center">无符合条件的记录！</div></td>
              </tr>
            </table>
			<% end if %>
          </div></td>
        <td width="10">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
<table width="610" height="60" cellpadding="-2" cellspacing="-2" background="../Images/manage_03.gif">
  <tr>
    <td width="608">&nbsp;</td>
  </tr>
</table>
</body>
</html>
