<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"-->
<%
Set rs=Server.CreateObject("ADODB.Recordset")
Sql="Select * from tb_classresource"
rs.open Sql,conn,1,3
rs.pagesize=5
'实现分页
if rs.Eof then
	rs_total = 0
else
   rs_total = rs.RecordCount
end if
dim pageno
getpageno = (Request("pageno"))
if(getpageno = "")then
	pageno = 1
else
	pageno = getpageno
End if
if(not rs.Eof)then
	rs.AbsolutePage = pageno
%> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>分页显示</title>
<style type="text/css">
<!--
body {
	margin-top: 0px;
	margin-bottom: 0px;
}
.style2 {font-size: 9pt}
.style4 {font-size: 9pt; color: #0C5C95; }
a{
color: #0C5C95;
 }
.style6 {
	font-size: 10pt;
	color: #0C5C95;
	font-family: "宋体";
}
.style8 {font-size: 10pt; color: #0C5C95; font-family: Arial, Helvetica, sans-serif; }
-->
</style></head>

<body>
<!--#include file="top.asp"-->
      <table width="882" height="218" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="21" height="218" valign="top"><img src="images/index_03.gif" width="21" height="79"></td>
          <td valign="top"><table width="801" height="88" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td height="31" valign="top"><span class="style4">　当前位置--&gt;&gt;<a href="index.asp" style="color:#0C5C95;text-decoration:none;">首页</a>--&gt;&gt;课程资讯---课程列表</span></td>
            </tr>
            <tr>
              <td height="57" valign="middle">　　　
                <table width="387" height="25" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="49" height="25">&nbsp;</td>
                    <td width="338" valign="top" background="images/index_35.gif"> 　　　<span class="style6">课程列表</span> 　　<span class="style8">　<em>class list </em></span></td>
                  </tr>
                </table></td>
            </tr>
          </table>                        <table width="750" height="41" border="1" align="center" cellpadding="0" cellspacing="0"  bordercolorlight="#A4BEE9" bordercolordark="#ffffff">
              <tr>
                <td width="50" height="11" valign="bottom"><div align="center"><span class="style4">课程编号</span></div></td>
                <td width="228" height="11" valign="bottom"><div align="center"><span class="style4">课程名称</span></div></td>
                <td width="107" height="11" valign="bottom"><div align="center"><span class="style4">开课时间</span></div></td>
                <td width="102" height="11" valign="bottom"><div align="center"><span class="style4">学时</span></div></td>
                <td width="123" height="11" valign="bottom"><div align="center"><span class="style4">学费</span></div></td>
                <td width="126" height="11" valign="bottom" class="style4"><div align="center">课程类型</div></td>
            </tr>
              <%
	         repeat_rows = 0
	         While(repeat_rows < rs.PageSize and Not rs.Eof)
	          %>
              <tr style="cursor:hand" onMouseOver="this.style.background='#eeeeee'" onMouseOut="this.style.background='#ffffff'">
                <td height="28" align="center" valign="middle" class="style4"><div align="center"><%=server.htmlencode(rs("class_id"))%>&nbsp;</div></td>
                <td height="28" align="center" valign="middle" class="style4">
				<div align="left">
				<%
				If(Len(rs("class_name")) > 30)Then 
				Response.Write("<a href='class_xianxi.asp?id="&rs("class_id")&"'>"&(Mid(rs("class_name"),1,30))&"...</a>")
				Else
				Response.Write("<a href='class_xianxi.asp?id="&rs("class_id")&"'>"&(rs("class_name"))&"</a>")
				End If
			    %>
				</div>
				</td>
                <td height="28" align="center" valign="middle" class="style4"><div align="center"><%=server.htmlencode(rs("class_time"))%>&nbsp;</div></td>
                <td height="28" align="center" valign="middle" class="style4"><div align="center"><%=server.htmlencode(rs("class_hour"))%>&nbsp;</div></td>
                <td height="28" align="center" valign="middle" class="style4"><div align="center"><%=server.htmlencode(rs("class_money"))%>&nbsp;元</div></td>
                <td height="28" align="center" valign="middle" class="style4"><div align="center"><%=server.htmlencode(rs("class_type"))%>&nbsp;</div></td>
              </tr>
              <%
	          repeat_rows = repeat_rows + 1
	          rs.MoveNext
	          Wend
	          %>
            </table>
            <table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td width="321" height="27" class="style4">[<%=pageno%>/<%=rs.PageCount%>]&nbsp;每页<%=rs.PageSize%>条&nbsp;共<%=rs_total%>条记录 </td>
                <td width="476">
                  <div align="right" class="style4">
                    <%
					if(pageno <> 1)then
					%>
                    <a href="?">第一页</a>
                    <%
					End if
					if(pageno <> 1)then
					%>
                    <a href="?pageno=<%=(pageno-1)%>">上一页</a>
                    <%
					end if
					if(instr(pageno,cstr(rs.pagecount)) = 0)then
					%>
                    <a href="?pageno=<%=(pageno+1)%>">下一页</a>
                    <%
					end if
					if(instr(pageno,cstr(rs.pagecount)) = 0)then
					%>
                    <a href="?pageno=<%=rs.pagecount%>">最后一页</a>
                    <%
					end if
					rs.close
					Set rs = Nothing
					end if
					%>
                </div></td>
              </tr>
            </table>
          <p>&nbsp;</p></td>
          <td width="21" valign="top"><img src="images/index_04.gif" width="21" height="79"></td>
        </tr>
      </table>
      <table width="882" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><!--#include file="down.asp"--></td>
        </tr>
      </table>
</body>
</html>
