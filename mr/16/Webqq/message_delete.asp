<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
set conn=server.CreateObject("adodb.connection")
conn.open "driver={SQL Server};server=.;uid=sa;pwd=;database=db_QQ"
set rs=server.CreateObject("adodb.recordset")
sql="select * from tb_talk"
rs.open sql,conn,1,3
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>清除所有聊天记录信息</title>
<style type="text/css">
<!--
.style1 {color: #FFFFFF}
.STYLE3 {color: #FFFFFF; font-size: 9pt; }
.STYLE7 {font-size: 9pt; color: #000000; }
.STYLE9 {
	font-size: 9pt;
	color: #FF0000;
}
-->
</style>
</head>
<body>
 <p align="center">
      <input name="po" type="submit" id="po" value="清除所有聊天记录" onClick="if( confirm('是否真的删除所有聊天记录?')){window.location.href='delete.asp';}">
</p>
 <form name="form1" method="post" action="message_delete.asp">
  <div align="center">
   
    <span class="STYLE9">注意：在此只显示最新的前8条记录</span><br>
   <%
If(rs.Eof)Then
	Response.Write("暂无聊天信息!")
Else
%>
    <table width="503" border="1" cellspacing="0" bordercolorlight="#0000CC" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
      <tr align="center" bgcolor="#0000CC">
        <td width="23" bgcolor="#666666"><span class="STYLE3">ID</span></td>
        <td width="66" bgcolor="#666666"><span class="STYLE3">发送者的ID</span></td>
        <td width="68" bgcolor="#666666"><span class="STYLE3">发送者昵称</span></td>
        <td width="49" bgcolor="#666666" class="STYLE3">接收者ID</td>
        <td width="106" bgcolor="#666666" class="STYLE3">信息内容</td>
        <td width="76" bgcolor="#666666" class="STYLE3">发送日期</td>
        <td width="53" bgcolor="#666666" class="STYLE3">回复标识</td>
        <td width="28" bgcolor="#666666" class="STYLE3">操作</td>
      </tr>
 <%
		'分页
	  rs.pagesize=6
	  page1=clng(request("page1"))
	  if page1<1 then page1=1
	  rs.absolutepage=page1
	  for i=1 to rs.pagesize
	  k=rs("message")
%>
      <tr align="center">
        <td height="22"><span class="STYLE7"><%=rs("id")%></span></td>
        <td><span class="STYLE7"><%=rs("fuserid")%></span></td>
        <td><span class="STYLE7"><%=rs("fusernickname")%></span></td>
        <td><span class="STYLE7"><%=rs("touserid")%></span></td>
        <td><span class="STYLE7">
		<% if len(k)>5 then%>
       <%=replace(left(k,5)&"....",chr(13),"<br>")%>
       <%else%>
        <%=k%>
        <%end if %>
</span></td>
        <td><span class="STYLE7"><%=rs("senddate")%></span></td>
        <td><span class="STYLE7"><%=rs("isreaded")%></span></td>
        <td><span class="STYLE7"><a href="#" onClick="if(confirm('是否确认删除?')){window.location.href='messagedel_deal.asp?id=<%=rs("id")%>';}">删除</a></span></td>
      </tr>
      <%
rs.movenext
if rs.eof then exit for
next
%>
    </table><br>
    <table width="503" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td><div align="right" class="STYLE7">
                    <% if page1<>1 then%>
                    <a href="<%=path1%>?page1=1">第一页</a> <a href="<%=path1%>?page1=<%=(page1-1)%>"> 上一页</a>
                    <%end if %>
                    <%if page1<>rs.pagecount then%>
                    <a href="<%=path1%>?page1=<%=(page1+1)%>">下一页</a> <a href="<%=path1%>?page1=<%=rs.pagecount%>" >最后一页</a>
                    <%end if%>
                </div></td>
              </tr>
    </table> 
    <p><%end if%></p>
</div>
</form>
</body>
</html>
