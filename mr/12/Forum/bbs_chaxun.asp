<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<%
'在显示留言信息时，应用了视图view_Board，该视图中涉及了表tb_User、tb_Topic、tb_Reply、tb_smalltype
If Request("keyword")<>"" Then
	keyword=Request("keyword")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	SQL="Select * From view_Board Where subject like '%"&keyword&"%' or content like '%"&keyword&"%' Order By CreateTime DESC"
	rs.Open SQL,Conn,1,3
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>查询页面</title>
<style type="text/css">
<!--
.STYLE2 {font-size: 9pt}
-->
</style>
<style type="text/css">
<!--
.STYLE2 {font-size: 9pt}
a:link {
	text-decoration: none;
	color: #000000;
}
a:visited {
	text-decoration: none;
	color: #000000;
}
a:hover {
	text-decoration: none;
	color: #000000;
}
a:active {
	text-decoration: none;
	color: #000000;
}
.STYLE3 {font-size: 9pt}
.STYLE3 {font-size: 9pt}
-->
</style>
</head>
<script language="javascript">
function mycheck()
{
if(findbbs.keyword.value=="")
{alert("请输入所要查询的关键字!");findbbs.keyword.focus();return false}
findbbs.submit();
}
</script>
<body>
<link rel="stylesheet" type="text/css" href="css/style.css">
<!--#include file="top.asp"-->
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" height="23" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td bgcolor="#E9EBEB"><table width="100%" height="22" border="0" align="center" cellpadding="0" cellspacing="0">
         <form name="findbbs" method="post" action="bbs_chaxun.asp" onSubmit="return mycheck(this)">
		  <tr>
            <td width="3%"><div align="right"><img src="images/bbs_biao.gif" width="16" height="14"></div></td>
            <td width="61%">&nbsp;<span class="STYLE2">编程词典技术论坛&nbsp;》&nbsp;查找帖子</span></td>
            <td width="20%"><input name="keyword" type="text" class="inputcss1" id="keyword" size="30"></td>
            <td width="1%">&nbsp;</td>
            <td width="15%"><input type="image"  src="images/findbbs_button.gif" alt="查找帖子"></td>
          </tr>
		  </form>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table><br>
<%if rs.eof then%>
<table width="96%" height="152"  border="0" align="center" cellpadding="-2" cellspacing="-2">
<tr>
<td width="15%" rowspan="2" align="center"><img src="Images/clue.gif" width="50" height="57"></td>
<td width="82%" height="84"><br><br>查询关键字：" <%=Request("keyword")%> "</td>
<td width="3%" rowspan="3">&nbsp;</td>
</tr>
<tr>
<td height="34" style="color:#660033 ">很报歉！没有找到任何您要查询的信息，请您用其他关键字再试一次！</td>
</tr>
<tr>
<td height="34" colspan="2" align="center"><div align="center">[ <a href="main.asp">返回</a> ]</div></td>
</tr>
</table>
<%else%>
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#349ED8">
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="1" align="center" cellpadding="0" cellspacing="1" bordercolor="#349ED8">
      <tr>
        <td width="54%" height="22" bgcolor="#58BAE9" >&nbsp;&nbsp;<font color="#FFFFFF" class="STYLE3">主题</font></td>
        <td width="15%" bgcolor="#58BAE9" ><div align="center" class="STYLE3"><font color="#FFFFFF">发贴人</font></div></td>
        <td width="20%" height="22" bgcolor="#58BAE9" ><div align="center" class="STYLE3"><font color="#FFFFFF">发帖时间</font></div></td>
        <td width="18%" bgcolor="#58BAE9" ><div align="center" class="STYLE3">回复|人气</div></td>
 </tr>
<%
rs.pagesize=10
page=CLng(Request("page"))
if page<1 then page=1
rs.absolutepage=page
for i=1 to rs.pagesize
aa=rs("subject")
zz=replace(aa,keyword,"<font color='#FF0000'>"&keyword&"</font>")
%>
<tr>
<td height="22">&nbsp;&nbsp;<a href="show.asp?TopicID=<%=rs("ID")%>"><%=zz%> &nbsp;&nbsp;</a></td>
        <td><div align="center"><%=rs("username")%></div></td>
        <td height="22"><div align="center"><%=rs("createtime")%></div></td>
        <td height="22"><div align="center">&nbsp;<%=rs("reply")%>/<%=rs("hit")%></div></td>
      </tr>
      
<%
m=m+1  '注意:该条语句一定不能少
rs.movenext
If rs.Eof Then exit For 
next
%>	
	  <tr>
        <td height="22" colspan="4" style="line-height:2"><table width="281" height="24" border="0" align="right" cellpadding="0" cellspacing="0">
          <tr>
            <td>
			  <div align="right">
			    <% if page<>1 then %>
			    <a href=?page=1&keyword=<%=keyword%>>第一页</a>
			    <a href=?page=<%=(page-1)%>&keyword=<%=keyword%>>上一页</a> 
			        <%
			end if 
			if page<>rs.pagecount then %>
			    <a href=?page=<%=(page+1)%>&keyword=<%=keyword%> >下一页</a> 
			    <a href=?page=<%=rs.pagecount%>&keyword=<%=keyword%>>最后一页</a> 
		            <%end if%>	
			      </div></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>
<%end if%>
</body>
</html>
