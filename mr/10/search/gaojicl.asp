<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn/conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>搜索引擎</title>
<style>
td{
font-size:12px;}
.STYLE1 {font-size: 9px}
.STYLE2 {font-size: 9pt; }
</style>
</head>
<%
if request.form("select1")<>"" then
	inpage=request.form("select1")
else
	inpage=request.querystring("select1")
	if inpage="" then
		inpage=10
	end if
end if
if request.form("content")<>"" then
	content=request.form("content")
else
	content=request.querystring("content")
end if 
if request.form("type1")<>"" then
	type1=request.form("type1")
else
	type1=request.querystring("type1")
end if 
set rs=server.CreateObject("adodb.recordset")
rs1="select * from tb_sousuo  where name1 like '%"&name1&"%' and content like '%"&content&"%' and type1 like '%"&type1&"%'"
rs.open rs1,conn,1,1 
%>
<script language="javascript">
function Mycheck(){
if (form1.name1.value=="")
{ alert("请输入要搜索的全部名称！");form1.name1.focus();return;}
form1.submit();
}
</script>
<body onload="clockon()">
<!--#include file="shangbiao.asp"-->
<table width="780" height="160" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="757" height="70" valign="top">
      <p align="right"><a href="gaojisousuo.asp" class="STYLE2"style="color:#3661a6;text-decoration:none">［返回］</a> </p>
      <table width="364" height="39" border="0" align="center">
          <tr>
            <td height="35" valign="top"><span class="STYLE2">
            <%if rs.eof and rs.bof then %>
            <img src="images/turnbook.gif" width="37" height="30">
            <% response.write "对不起，没有找到您要搜索的信息！"
				  response.end
				  else
				  if content="" or type1="" then%> 
			    <img src="images/turnbook.gif" width="37" height="30"> 
			    <% response.write "对不起，没有找到您要搜索的信息！"
				  response.end
				  else%>
	        <%'分页
			rs.pagesize=clng(inpage)
			page=clng(request.querystring("page"))
			if page<1 then page=1
			rs.absolutepage=page
			for i=1 to rs.pagesize
			jjs=rs("content")
            jj=replace(jjs,content,"<font color='#FF0000'>"&content&"</font>")
			%>			
	        </span></td>
        </tr> 
      </table>    </td>
  <tr>
    <td height="54" valign="top"><table width="658" height="54" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="769" height="54" valign="top"><span class="STYLE2"><a href="#"onClick="javascript:window.open('introduce9.asp?id=<%=rs("ID")%>')"><%=rs("name1")%></a><br>
			
              <% if len(jj)>150 then%>
		      <%=left(jj,150)&"...."%>
              <%else%>
              <%=jj%>
              <%end if %>
              <br>
              <%=rs("type1")%></span></td>
          </tr> <%
		  rs.movenext
		  if rs.eof then exit for
 next
 %>
</table></td></tr>
  <tr>
    <td height="9" valign="top">
 <div align="left" class="STYLE2">
  <% if page<>1 then %>
   <a href=<%=path%>?page=1&inpage=<%=inpage%>&content=<%=content%>&name1=<%=name1%>&type1=<%=type1%>>第一页</a> <a href=<%=path%>?page=<%=(page-1)%>&inpage=<%=inpage%>&content=<%=content%>&name1=<%=name1%>&type1=<%=type1%>>上一页</a>
   <%end if 
	if page<>rs.pagecount then %>
   <a href=<%=path%>?page=<%=(page+1)%>&inpage=<%=inpage%>&content=<%=content%>&name1=<%=name1%>&type1=<%=type1%>>下一页</a> <a href=<%=path%>?page=<%=rs.pagecount%>&inpage=<%=inpage%>&name1=<%=name1%>&type1=<%=type1%>&content=<%=content%>>最后一页</a>
   <%end if %>
</div></td>
</tr>
  <tr>
    <td height="9" valign="top"><!--#include file="copyright.asp"-->&nbsp;</td>
  </tr>
</table>
<%end if 
end if %>
</body>
</html>
