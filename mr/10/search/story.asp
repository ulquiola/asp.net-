<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="shangbiao.asp"-->
<!--#include file="Conn/conn.asp"-->
<%
set rs=server.CreateObject("adodb.recordset")
rs_1="select * from tb_story"
rs.open rs_1,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<body onLoad="clockon()" background="images/bg.gif">
<table width="780" height="401" border="0" align="center" cellpadding="０" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td width="779" valign="top"><table width="780" height="51" border="0" cellpadding="0" cellspacing="0">
        <tr valign="top">
          <td width="20" background="images/sousou2.gif">&nbsp;</td>
          <td width="76" height="47" background="images/tubiao.gif">&nbsp;</td>
          <td width="4" height="47" background="images/sousou2.gif" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="80" background="images/chuangyegushi.gif" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="381" align="right" background="images/sou2.gif" bgcolor="#FFFFFF"><a href="index.asp"style="color:#3661a6"></a></td>
          <td width="73" align="right" valign="bottom" background="images/sou2.gif" bgcolor="#FFFFFF"><div align="center"><a href="index.asp"style="color:#3661a6; font-size: 9pt">返回首页</a></div></td>
          <td width="91" background="images/sousouhou.gif" bgcolor="#FFFFFF"><a href="index.asp"style="color:#3661a6"></a></td>
        </tr>
        <tr>
          <td height="3" colspan="7" background="images/erjiye1.gif"></td>
        </tr>
        <tr>
          <td height="30" colspan="7">&nbsp;</td>
        </tr>
      </table></p>
    <table width="701" height="50" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#ccccff" style="color:#336699 ">
        <%'分页
	  rs.pagesize=12
	  page=clng(request("page"))
	  if page<1 then page=1
	  rs.absolutepage=page
	  for i=1 to rs.pagesize
	  k=rs("content")
	  %>
        <tr>
          <td><p><span style="font-size: 9pt"><a href="#"onclick="javascript:window.open('introduce2.asp?id=<%=rs("id")%>')"><%=rs("title")%></a><br>
              <% if len(k)>200 then%>
               <%=replace(left(k,200)&"....",chr(13),"<br>")%>
              <%else%> 
              <%=k%>
              <%end if %>
          </span><br> 
          <hr color="#ffffff"></hr>
	      </td>
        </tr>
        <%
	  rs.movenext
	  if rs.eof then exit for 
	  next
	  %>
      </table>      <table width="701" height="20" border="0" align="center">
        <tr>
          <td> <div align="right" style="font-size: 9pt">
              <% if page<>1 then%>
              <a href=<%=path%>?page=1>第一页</a> <a href=<%=path%>?page=<%=(page-1)%>> 上一页</a>
              <%end if %>
              <%if page<>rs.pagecount then%>
              <a href=<%=path%>?page=<%=(page+1)%>>下一页</a> <a href=<%=path%>?page=<%=rs.pagecount%> >最后一页</a>
              <%end if%>
              </div></td>
        </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
