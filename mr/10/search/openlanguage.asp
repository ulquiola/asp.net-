<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="shangbiao.asp"-->
<!--#include file="Conn/conn.asp"-->
<%
set ab=server.CreateObject("adodb.recordset")
ab_1="select * from tb_open"
ab.open ab_1,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>明日在线搜索</title>
</head>
<body onLoad="clockon()"background="images/bg.gif">
<tr>
    <td height="405"></td>
  <td></td>
</tr>
<table width="780" height="421" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" >
  <tr align="right">
    <td width="39" align="right" ></td>
    <td width="740" align="left" valign="top"><table width="780" height="51" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="21" background="images/sousou2.gif">&nbsp;</td>
          <td width="81" height="47" background="images/tubiao.gif">&nbsp;</td>
          <td width="7" height="47" valign="bottom" background="images/sousou2.gif" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="117" valign="bottom" background="images/kaifayuyanjianjie.gif" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="376" align="right" valign="bottom" background="images/sou2.gif" bgcolor="#FFFFFF"><a href="index.asp"style="color:#3661a6"></a></td>
          <td width="78" align="right" valign="bottom" background="images/sou2.gif" bgcolor="#FFFFFF"><div align="center"><a href="index.asp"style="color:#3661a6; font-size: 9pt">返回首页</a></div></td>
          <td width="100" valign="bottom" background="images/sousouhou.gif" bgcolor="#FFFFFF"><a href="index.asp"style="color:#3661a6"></a></td>
        </tr>
        <tr>
          <td height="3" colspan="7" background="images/erjiye1.gif"></td>
        </tr>
        <tr>
          <td height="30" colspan="7">&nbsp;</td>
        </tr>
      </table>
      <table width="696" height="24" border="0" align="center">
        <tr> <%'分页
	ab.pagesize=12
	page=clng(request("page"))
	if page<1 then page=1
	ab.absolutepage=page
	for i=1 to ab.pagesize
	zz=ab("introduce")
	%>
          <td width="690"><table width="691" height="69" border="0" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bordercolordark="#ccccff" bordercolorlight="#FFFFFF"style="color:#336699">
            <tr>
              <td width="691"><p><span style="font-size: 9pt"><a href="#"onClick="javascript:window.open('introduce5.asp?id=<%=ab("ID")%>','')"><%=ab("name1")%></a>
                    <br>
			        <% if len(zz)>200 then%>
                    <%=replace(left(zz,200)&"....",chr(13),"<br>")%>
			        <%else%>
			        <%=zz%>
			        <%end if %>
              </span><br>
					<hr color="#FFFFFF" size="1"></hr>
              </td>
              </tr>
            <%
	  ab.movenext
	  if ab.eof then exit for 
	  next
	  %>
          </table>
            <table width="691" height="24" border="0">
              <tr>
                <td width="685" height="5" align="right">
                                      <div align="right" style="font-size: 9pt">
                                          <% if page<>1 then%>
                                          <a href=<%=path%>?page=1>第一页</a> <a href=<%=path%>?page=<%=(page-1)%>> 上一页</a>
                                          <%end if %>
                                          <%if page<>ab.pagecount then%>
                                          <a href=<%=path%>?page=<%=(page+1)%>>下一页</a> <a href=<%=path%>?page=<%=ab.pagecount%> >最后一页</a>
                                          <%end if%>
                </div></td>
              </tr>
            </table>
         </td>
        </tr>
    </table>    </td>
  </tr>
</table>
</body>
</html>
