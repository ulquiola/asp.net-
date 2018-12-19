<!--#include file="conn.asp" -->
<%
Set rs1=Server.CreateObject("ADODB.RecordSet")			'创建记录集
SQL="Select * From tb_wish"								'查询数据
rs1.Open SQL,conn,1,3									'打开记录集
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<style><!--@import url(CSS/list.css);--></style>
<script type="text/javascript" src="inc/common.js"></script>
</head>
<body>
<div id="header">&nbsp;</div>
<div id="menu">
  <a href="index.asp"><img src="images/btn_index.gif" width="87" height="22" border="0" /></a>
  <a href="message_list.asp"><img src="images/btn_list.gif" alt="字条列表" border="0" /></a>
  <a href="message_list.asp?flag=lei"><img src="images/rank.gif" alt="帖排行" border="0" /></a>
</div>
<%
if rs1.eof then
	response.Write("暂无祝福信息！")
else
%>
<%end if%>
<div id="mainList">
	<table width="100%" border="0" cellpadding="0" cellspacing="1" class="List">
		<tr>
			<th>编号</th>
			<th>字条内容</th>
			<th>祝福对象</th>
			<th>祝福者</th>
			<th>人气</th>
			<th>发送时间</th>
		</tr>
	<%
	Dim Sql
	Set rs=Server.CreateObject("ADODB.RecordSet")			'创建记录集
	If request("flag")="lei" then
	Sql="Select * From tb_wish Order By hits DESC"
	Else
		If request("ID")="" Then
			Sql="Select * From tb_wish Order By hits DESC"			
		Else  
			Dim ID
			ID=CLng(Trim(request("ID")))
			Sql="Select * From tb_wish Where ID="&ID
		End If 
	End If
	rs.Open SQL,conn,1,3
		'分页显示用户信息
		rs.pagesize=10								'设置分页显示时，每页显示记录的数目
		page1=CLng(Request("page1"))				'将获取到的记录数转换为整数
		if page1<1 then page1=1						'为page变量设置默认值
		rs.absolutepage=page1						'在进行分页显示时，指定指针所在的页
		for i=1 to rs.pagesize						'循环显示记录信息
		%>
		<tr>
			<td class="ListID"><%=Rs("ID")%></td>
			<td class="ListInfo"><%=Rs("Content")%></td>
			<td class="ListSend"><%=Rs("wishMan")%></td>
			<td class="ListPick"><%=Rs("wellWisher")%></td>
			<td class="Listhits"><%=Rs("hits")%></td>
			<td class="ListDate"><%=Rs("sendTime")%></td>
		</tr>
	<%
		Rs.MoveNext 
		If rs.Eof Then exit For 		'退出For循环
		Next
	%>   
	</table>
	 <table width="100%" height="26"  border="0" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
	<%If not(rs.Eof and rs.Bof) Then%>	
    <td width="34%" align="left" bgcolor="#FFFFFF">&nbsp;
	      		<% if page1<>1 then %><a href=?page1=1 class="white">第一页</a>
				<a href=?page1=<%=(page1-1)%> class="white">上一页</a> 
			<%
			end if 
			if page1<>rs.pagecount then %>
   				<a href=?page1=<%=(page1+1)%> class="white">下一页</a> 
				<a href=?page1=<%=rs.pagecount%> class="white">最后一页</a> 
	  <%end if%>	  </td>
    <td width="66%" align="right" bgcolor="#FFFFFF" class="word_grey">[<%=page1%>/<%=rs.PageCount%>]&nbsp;&nbsp;每页<%=rs.PageSize%>条&nbsp;&nbsp;共<%=rs.RecordCount%>条字条&nbsp;</td>
		<%End If%>	
  </table>
</div>
<div id="footer">&nbsp;</div>
<%=Bottom%>
</body>
</html>
