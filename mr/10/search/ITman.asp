<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"-->
<%
set rs=server.CreateObject("adodb.recordset")
rs_user="select * from tb_ITman where type1='国内'"
rs.open rs_user,conn,1,3
set rd=server.createobject("adodb.recordset")
rrd="select * from tb_ITman where type1='国外'"
rd.open rrd,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>明日在线搜索</title>
<style type="text/css">
<!--
body {
	margin-top: 0px;
}
.style1 {font-size: 9px}
.style2 {font-size: 9pt}
.STYLE3 {font-size: 10pt; }
-->
</style></head>
<body  onLoad="clockon()" background="images/bg.gif">
<!--#include file="shangbiao.asp"-->
<table width="780" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" algin="center" > 
<tr>
  <td width="780" height="681"></td>
	<td width="780" valign="top"><table width="780" height="668" border="0" cellpadding="0" cellspacing="0" align="center">
      <tr>
        <td height="668" align="center" valign="top"><table width="707" border="0" align="center" cellpadding="0" cellspacing="0" >
              <tr>
                <td height="447" valign="top"><div align="center">
                  <table width="780" height="51" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="21" background="images/sousou2.gif">&nbsp;</td>
          <td width="81" height="47" background="images/tubiao.gif">&nbsp;</td>
          <td width="7" height="47" valign="bottom" background="images/sousou2.gif" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="139" valign="bottom" background="images/ITfongyunrenwujianjie.gif" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="354" align="right" valign="bottom" background="images/sou2.gif" bgcolor="#FFFFFF"><a href="index.asp"style="color:#3661a6"></a></td>
          <td width="78" align="right" valign="bottom" background="images/sou2.gif" bgcolor="#FFFFFF"><div align="center"><a href="index.asp" class="style2"style="color:#3661a6">返回首页</a></div></td>
          <td width="100" valign="bottom" background="images/sousouhou.gif" bgcolor="#FFFFFF"><a href="index.asp"style="color:#3661a6"></a></td>
        </tr>
        <tr>
          <td height="3" colspan="7" background="images/erjiye1.gif"></td>
        </tr>
        <tr>
          <td height="30" colspan="7">&nbsp;</td>
        </tr>
      </table>
                  
                    <div align="center">
                      <table width="675" height="82" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#ccccff" style="color:#336699 ">
                        <tr>
                          <td align="center"><div align="center" class="STYLE3">国内IT风云人物简介</div></td>
                        </tr>
                        <tr>
                          <td></td>
                        </tr>
                        <% '分页
		    rs.pagesize=6
			page=clng(request("page"))
			if page<1 then page=1
			rs.absolutepage=page
			for i=1 to rs.pagesize
			jk=rs("introduce")
			%>
                        <tr>
                          <td><span class="style2"><a href="#"onClick="javascript:window.open('introduce1.asp?id=<%=rs("ID")%>')"><%=rs("name1")%></a><br>
                            <% if len(jk)>200 then%>
                            <%=replace(left(jk,200)&"....",chr(13),"<br>")%>
                            <%else%>
                            <%=jk%>
                            <%end if %>
                          </span><br>
                          <hr color="#ffffff"></hr>						</td>
                   <%           
			rs.movenext
			if rs.eof then exit for
			next
			%>
                      </table>
                      <table width="675" border="0">
                        <tr>
                          <td width="604" height="15">
                            <div align="right" class="style2">
                              <% if page<>1 then%>
                              <a href=<%=path%>?page=1&page1=<%=clng(page1)%>>第一页</a> <a href=<%=path%>?page=<%=(page-1)%>&page1=<%=clng(page1)%>> 上一页</a> <%end if %>
                              <%if page<>rs.pagecount then%>
                              <a href=<%=path%>?page=<%=(page+1)%>&page1=<%=clng(page1)%>>下一页</a> <a href=<%=path%>?page=<%=rs.pagecount%>&page1=<%=clng(page1)%>>最后一页</a>
                              <%end if%>
                          </div></td>
                        </tr>
                        </table>
                      <table width="675" border="0">
                        <tr>
                          <td><img src="images/FGline.gif"></td>
                        </tr>
                        </table>
                      <table width="675" height="82" border="0" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bordercolordark="#ccccff" bordercolorlight="#FFFFFF"style="color:#336699 ">
                        <tr>
                          <td><div align="center" class="style2">国外IT风云人物简介</div></td>
                        </tr>
                        <tr>
                          <td>&nbsp;</td>
                        </tr>
                        <%
			'分页
			rd.pagesize=6
			page1=clng(request("page1"))
			if page1<1 then page1=1
			rd.absolutepage=page1
			for i=1 to rd.pagesize
			aa=rd("introduce")
			%>
                        <tr>
                          <td><span class="style2"><a href="#"onClick="javascript:window.open('introduce1.asp?id=<%=rd("ID")%>')"><%=rd("name1")%></a><br>
                            <% if len(aa)>200 then%>
                            <%=left(aa,200)&"...."%>
                            <%else%>
                            <%=aa%>
                            <%end if %>
                          </span><br>
                          <hr color="#ffffff"></hr>						</td>
                          <%
			rd.movenext
			if rd.eof then exit for
			next
			%>
                      </table>
                      <table width="684" border="0" align="center">
                        <tr>
                          <td width="678" align="center">
                            <div align="right">
                              <span class="style2">
                              <% if page1<>1 then %>
                              <a href=<%=path%>?page1=1&page=<%=clng(page)%>>第一页</a> <a href=<%=path%>?page1=<%=(page1-1)%>&page=<%=clng(page)%>>上一页</a>
                              <%end if 
		if page1<>rd.pagecount then %>
                              <a href=<%=path%>?page1=<%=(page1+1)%>&page=<%=clng(page)%>>下一页</a> <a href=<%=path%>?page1=<%=rd.pagecount%>&page=<%=clng(page)%>>最后一页</a>
                              <%end if %>
                              </span>
                              <p></p>
                          </div></td>
                        </tr>
                        </table>
                    </div>
                </div>                </td>
              </tr>
          </table></td>
        </table>
 </td>
</tr>
</table>
</body>
</html>
