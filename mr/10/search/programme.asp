<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="shangbiao.asp"-->
<!--#include file="Conn/conn.asp"-->
<%
set ac=server.CreateObject("adodb.recordset")
bc="select * from tb_programme"
ac.open bc,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>明日在线搜索</title>
</head>
<body onload="clockon()" background="images/bg.gif">
<table width="779" height="350" border="0" align="center" cellpadding="00" cellspacing="0" bgcolor="#FFFFFF">
  <tr valign="top">
    <td width="807" height="350"><table width="780" height="51" border="0" cellpadding="0" cellspacing="0">
        <tr valign="top">
          <td width="20" background="images/sousou2.gif">&nbsp;</td>
          <td width="76" height="47" background="images/tubiao.gif">&nbsp;</td>
          <td width="4" height="47" background="images/sousou2.gif" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="80" background="images/bianchengwangzhan.gif" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="381" align="right" background="images/sou2.gif" bgcolor="#FFFFFF"><a href="index.asp"style="color:#3661a6"></a></td>
          <td width="73" align="right" background="images/sou2.gif" bgcolor="#FFFFFF"><div align="center"><a href="index.asp"style="color:#3661a6; font-size: 9pt">返回首页</a></div></td>
          <td width="91" background="images/sousouhou.gif" bgcolor="#FFFFFF"><a href="index.asp"style="color:#3661a6"></a></td>
        </tr>
        <tr>
          <td height="3" colspan="7" background="images/erjiye1.gif"></td>
        </tr>
        <tr>
          <td height="30" colspan="7">&nbsp;</td>
        </tr>
      </table></p>
    <table width="702" height="163" border="0" align="center">
        <tr>
          <td height="159"><table width="702" height="29" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#ccccff"style="color:#336699">
            <%'分页
	  ac.pagesize=10
	  page=clng(request("page"))
	  if page<1 then page=1
	  ac.absolutepage=page
	  for i=1 to ac.pagesize
	  jjk=ac("introduce")
	  %>
            <tr>
              <td><span style="font-size: 9pt"><a href="#"onClick="javascript:window.open('introduce6.asp?id=<%=ac("ID")%>')"><%=ac("name1")%></a> <br>
                  <% if len(jjk)>200 then%>
                  <%=replace(left(jjk,200)&"...."(13),"<br>")%>
                  <%else%>
                  <%=jjk%>
                  <%end if %>
                  <br>
                    <a href="<%=ac("netaddress")%>"><%=ac("netaddress")%></a></span><br>
				  <hr color="#ffffff"></hr>
			    </td>
            </tr>
            <%
	ac.movenext
	if ac.eof then exit for 
	next
	%>
          </table>
            <table width="702" border="0" align="center">
              <tr>
                <td width="531"><div align="right">
                    <span style="font-size: 9pt">
                    <% if page<>1 then%>
                    <a href=<%=path%>?page=1>第一页</a> <a href=<%=path%>?page=<%=(page-1)%>> 上一页</a>
                    <%end if %>
                    <%if page<>ac.pagecount then%>
                    <a href=<%=path%>?page=<%=(page+1)%>>下一页</a> <a href=<%=path%>?page=<%=ac.pagecount%> >最后一页</a>
                    <%end if%>
                    </span>
                    <p></p>
                </div></td>
              </tr>
            </table>
          <p>&nbsp;</p></td>
        </tr>
    </table>      </td>
  </tr>
</table>
</body>
</html>
