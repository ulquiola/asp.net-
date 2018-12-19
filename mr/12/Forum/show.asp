<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"-->
<%
If Session("UserName")<>"" Then
	UStatus=""
	Set rs=Server.CreateObject("ADODB.RecordSet")
	SQL="Select * From tb_User Where UserName='"&Session("UserName")&"'"
	rs.Open SQL,Conn,1,3	
	If rs.Eof and rs.Bof Then
		Session("UserName")=""
		UStatus=""
	Else
		If rs("Status")="管理员" or rs("Status")="版主" Then UStatus="管理员" End If
	End If
End If
'在显示主题信息时，应用了视图view_topic，该视图中涉及了表tb_User、tb_Topic
TopicID=Request.QueryString("TopicID")
If TopicID<>"" Then
	'更新帖子的点击数
	UPD="Update tb_Topic set hit=hit+1 where ID="&TopicID
	conn.execute UPD
	'查询主题信息
	Set rs=Server.CreateObject("ADODB.RecordSet")
	SQL="Select * from view_Topic where ID="&TopicID
	rs.Open SQL,Conn,1,3
Else
	response.Redirect("index.asp")
	response.End()
End If
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="./Css/style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
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
<body>
<%
'转换UBB代码
Function UBBCodeDeal(InString)
	InString1=Server.HTMLEncode(InString)
	InString1=Replace(InString1,"[","<")
	InString1=Replace(InString1,chr(13),"<BR>")
	UBBCodeDeal=Replace(InString1,"]",">")
End Function
%> 
<table width="777" height="341"  border="0" align="center" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
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
          <td width="13" height="32" background="Images/bg.gif">&nbsp;</td>
          <td width="689" background="Images/bg.gif" style="color:#FFFFFF"> <img src="Images/topic.gif" width="10" height="9">&nbsp;原主题：『<%=server.HTMLEncode(rs("subject"))%>』</td>
          <td width="73" background="Images/bg.gif">楼主</td>
        </tr>
        <tr valign="top">
          <td height="171" colspan="3"><table width="100%" height="129"  border="0" cellpadding="-2" cellspacing="-2">
            <tr>
              <td valign="top" align="center">
			  <%
			  If rs.Eof and rs.Bof Then
			  %>
					<table width="100%" height="152"  border="0" cellpadding="-2" cellspacing="-2">
                      <tr>
                        <td width="15%" align="center"><img src="Images/clue.gif" width="32" height="32"></td>
                        <td width="82%">很报歉！没有该主题信息！</td>
                        <td width="3%" rowspan="2">&nbsp;</td>
                      </tr>
                      <tr>
                        <td height="34" colspan="2" align="center"><div align="center">[ <a href="index.asp">查看主题分类</a> ]</div></td>
                      </tr>
                    </table>
					<%
			  Else%>
			  <table width="100%" height="109"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder_Top">
			  <tr>
			  <td height="5"></td>
			  </tr>
                <tr>
                  <td width="136" valign="top"><div align="center"><br>
				<%
				'获取用户信息
				ID=rs("ID")
				UserName=rs("UserName")
				Email=rs("Email")
				homepage=rs("homepage")
				QQ=rs("QQ")
				CreateTime=rs("CreateTime")
				Sex=rs("sex")
				ICO=rs("ICO")
				IP=rs("IP")
				IP=left(IP,InStrRev(IP,"."))&"…"
				%>
                     <%=UserName%></p>
                   <img src="Images/head/<%=ICO%>" width="60" height="60"><br>
                    我是：<%=Sex%>生<br>
                      心情：<img src="Images/Face/<%=rs("face")%>" width="20" height="20"> <br><br>
</div></td>
                  <td width="1" background="Images/line.gif"></td>
                  <td width="666" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
				  <script language="javascript">
				  	function del(ID){
					if(confirm("真的要删除该主题信息吗？")){
						window.location.href="Topic_del.asp?TopicID="+ID;
					 }
					}
				  </script>
                    <tr>
                      <td width="70%" height="30"> &nbsp;<img src="Images/email.gif" alt="<%=UserName%>的Email是：<%=Email%>" width="45" height="16">Email&nbsp; <img src="Images/home.gif" alt="<%=UserName%>的个人主页是：<%=homepage%>" width="16" height="16">个人主页&nbsp; <img src="Images/oicq.gif" alt="<%=UserName%>的QQ号码是：<%=QQ%>" width="16" height="16">QQ&nbsp; <img src="Images/ip.gif" alt="IP为：<%=IP%>" width="16" height="15">IP地址&nbsp; <img src="Images/time.gif" width="16" height="16">发表时间：<%=FormatDateTime(CreateTime,2)%></td>
                      <td width="30%">
					  
					  
					  <a href="Reply.asp?TopicID=<%=ID%>&state=1">引用</a>
					  <% If UStatus="版主" or UStatus="管理员" Then%>
					  <input type="image" class="noborder" onClick="del(<%=ID%>)" src="Images/del_1.gif" align="middle" border="0" onMouseMove="this.src='Images/del_2.gif'"  onMouseOut="this.src='Images/del_1.gif'">
					  <%else%>
					  删除
					  <%End If%>&nbsp; 
					  <a href="Reply.asp?TopicID=<%=TopicID%>"><img src="Images/reply.gif" width="15" height="15" border="0">回复</a>					  </td>
                    </tr>
                    <tr>
                      <td colspan="2"><hr align=left width="92%" size=1 noshade></td>
                    </tr>
<tr>
<%Content=rs("Content")%>
  <td colspan="2" style="padding-left:10px"><%=UBBCodeDeal(Content)%></td>
</tr>
                  </table></td>
                </tr>
              </table>
<%
set rs_reply=server.CreateObject("ADODB.RecordSet")
sql="select * from view_Reply where TopicID="&TopicID&" order by CreateTime asc"
rs_reply.open sql,conn,1,3
if not(rs_reply.eof and rs_reply.bof) Then%>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr><td><hr width="100%" size="1">
</td></tr>
</table> 
		<%
			'分页显示类别列表
			'rs_reply.pagesize=2
			rs_reply.pagesize=rs_reply.recordcount
			page=CLng(Request("page"))
			if page<1 then page=1
			rs_reply.absolutepage=page
			for i=1 to rs_reply.pagesize
			session("i")=i
			  %>
<table width="100%" height="109"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder_Top">
			  <tr>
			  <td height="25" colspan="3" bgcolor="88BAF9">&nbsp;&nbsp;&nbsp;<a name="<%response.Write(rs_reply("ID"))  '加入锚点用于定位回复记录%>"></a>回复主题：<%=server.HTMLEncode(rs_reply("subject"))%></td>
			  <td bgcolor="88BAF9">&nbsp;</td>
			  <td bgcolor="88BAF9">&nbsp;</td>
			  <td bgcolor="88BAF9"><div align="right"><%=i%>楼&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</div></td>
			  <td width="7" height="25" bgcolor="88BAF9"></td>
			  </tr>
                <tr>
                  <td width="136" valign="top"><div align="center"><br>
				<%
				'获取用户信息
				ID=rs_reply("ID")
				UserName=rs_reply("UserName")
				Email=rs_reply("Email")
				homepage=rs_reply("homepage")
				QQ=rs_reply("QQ")
				CreateTime=rs_reply("CreateTime")
				Sex=rs_reply("sex")
				ICO=rs_reply("ICO")
				IP=rs_reply("IP")
				IP=left(IP,InStrRev(IP,"."))&"…"
				%>
                <%=UserName%></p>
               <img src="Images/head/<%=ICO%>" width="60" height="60"><br>
                我是：<%=Sex%>生<br>
               心情：<img src="Images/Face/<%=rs("face")%>" width="20" height="20"> <br><br>
</div></td>
                  <td width="1" background="Images/line.gif"></td>
                  <td width="668" colspan="4" valign="top">
				  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
				  <script language="javascript">
				  	function del_R(ID){
					if(confirm("真的要删除该条回复信息吗？")){
						window.location.href="Reply_del.asp?ReplyID="+ID;
					 }
					}
				  </script>
                    <tr>
                      <td width="64%" height="30"> &nbsp;<img src="Images/email.gif" alt="<%=UserName%>的Email是：<%=Email%>" width="45" height="16">Email&nbsp; <img src="Images/home.gif" alt="<%=UserName%>的个人主页是：<%=homepage%>" width="16" height="16">个人主页&nbsp; <img src="Images/oicq.gif" alt="<%=UserName%>的QQ号码是：<%=QQ%>" width="16" height="16">QQ&nbsp; <img src="Images/ip.gif" alt="IP为：<%=IP%>" width="16" height="15">IP地址&nbsp; <img src="Images/time.gif" width="16" height="16">发表时间：<%=FormatDateTime(CreateTime,2)%></td>
                      <td width="36%" align="center">
					    <a href="Reply.asp?TopicID=<%=TopicID%>&replyID=<%=ID%>&state=0&i=<%=i%>">引用</a>
					  <% If UStatus="管理员" or UStatus="版主" Then%><input type="image" class="noborder" onClick="del_R(<%=ID%>)" src="Images/del_1.gif" align="middle" border="0" onMouseMove="this.src='Images/del_2.gif'"  onMouseOut="this.src='Images/del_1.gif'"><%End If%>
                      &nbsp; <a href="#"><img src="Images/top.gif" width="47" height="20" border="0"></a></td>
                    </tr>
                    <tr>
                      <td colspan="2"><hr align=left width="92%" size=1 noshade></td>
                    </tr>
<tr>
<%Content=rs_reply("Content")%>
  <td colspan="2" style="padding-left:10px"><%=UBBCodeDeal(Content)%></td>
</tr>
                  </table></td>
                </tr>
              </table>
			  <%
			  rs_reply.MoveNext
			  If rs_reply.Eof Then exit For 
			  Next
			  End If
end if
			  %>			  </td>
            </tr>
          </table>		  </td>
        </tr>
      </table>
<table width="100%" height="25"  border="0" cellpadding="-2" cellspacing="-2">
  <tr>
    <td width="3%" valign="bottom"></td>
	<%If not(rs_reply.Eof and rs_reply.Bof) Then%>	
    <td width="68%" align="left">
	      		<% if page<>1 then %><a href=?page=1&TopicID=<%=TopicID%> class="white">第一页</a>
				<a href=?page=<%=(page-1)%>&TopicID=<%=TopicID%> class="white">上一页</a> 
			<%end if 
			if page<>rs_reply.pagecount then %>
   				<a href=?page=<%=(page+1)%>&TopicID=<%=TopicID%> class="white">下一页</a> 
				<a href=?page=<%=rs_reply.pagecount%>&TopicID=<%=TopicID%> class="white">最后一页</a> 
		  <%end if%>	</td>
    <td width="29%" align="right" class="word_grey">[<%=page%>/<%=rs_reply.PageCount%>]&nbsp;&nbsp;每页<%=rs_reply.PageSize%>条&nbsp;&nbsp;共<%=rs_reply.RecordCount%>条回复信息</td>
		<%End If   
	%>	
  </tr>
</table>    </td>
  </tr>
  <tr>
    <td valign="top" background="Images/hr.gif"></td>
  </tr>
</table>
</body>
</html>
<%
rs.close()
Set rs=Nothing
%>
