<!--#include file="conn.asp" -->
<%
Set rs1=Server.CreateObject("ADODB.RecordSet")			'������¼��
SQL="Select * From tb_wish"								'��ѯ����
rs1.Open SQL,conn,1,3									'�򿪼�¼��
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
  <a href="message_list.asp"><img src="images/btn_list.gif" alt="�����б�" border="0" /></a>
  <a href="message_list.asp?flag=lei"><img src="images/rank.gif" alt="������" border="0" /></a>
</div>
<%
if rs1.eof then
	response.Write("����ף����Ϣ��")
else
%>
<%end if%>
<div id="mainList">
	<table width="100%" border="0" cellpadding="0" cellspacing="1" class="List">
		<tr>
			<th>���</th>
			<th>��������</th>
			<th>ף������</th>
			<th>ף����</th>
			<th>����</th>
			<th>����ʱ��</th>
		</tr>
	<%
	Dim Sql
	Set rs=Server.CreateObject("ADODB.RecordSet")			'������¼��
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
		'��ҳ��ʾ�û���Ϣ
		rs.pagesize=10								'���÷�ҳ��ʾʱ��ÿҳ��ʾ��¼����Ŀ
		page1=CLng(Request("page1"))				'����ȡ���ļ�¼��ת��Ϊ����
		if page1<1 then page1=1						'Ϊpage��������Ĭ��ֵ
		rs.absolutepage=page1						'�ڽ��з�ҳ��ʾʱ��ָ��ָ�����ڵ�ҳ
		for i=1 to rs.pagesize						'ѭ����ʾ��¼��Ϣ
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
		If rs.Eof Then exit For 		'�˳�Forѭ��
		Next
	%>   
	</table>
	 <table width="100%" height="26"  border="0" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
	<%If not(rs.Eof and rs.Bof) Then%>	
    <td width="34%" align="left" bgcolor="#FFFFFF">&nbsp;
	      		<% if page1<>1 then %><a href=?page1=1 class="white">��һҳ</a>
				<a href=?page1=<%=(page1-1)%> class="white">��һҳ</a> 
			<%
			end if 
			if page1<>rs.pagecount then %>
   				<a href=?page1=<%=(page1+1)%> class="white">��һҳ</a> 
				<a href=?page1=<%=rs.pagecount%> class="white">���һҳ</a> 
	  <%end if%>	  </td>
    <td width="66%" align="right" bgcolor="#FFFFFF" class="word_grey">[<%=page1%>/<%=rs.PageCount%>]&nbsp;&nbsp;ÿҳ<%=rs.PageSize%>��&nbsp;&nbsp;��<%=rs.RecordCount%>������&nbsp;</td>
		<%End If%>	
  </table>
</div>
<div id="footer">&nbsp;</div>
<%=Bottom%>
</body>
</html>
