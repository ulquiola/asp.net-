<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"--><!--数据库连接文件-->
<!--#include file="Replace.asp"--><!--数据处理文件-->
<%
Sql = "Select * from tb_news order by time1 Desc"'查询数据
Set rs = Server.CreateObject("ADODB.RecordSet")'创建记录集
rs.Open Sql,conn,1,3'打开记录集
rs.PageSize = 9		'设置每页显示记录的数目
'实现分页
if rs.Eof then		
	rs_total = 0	'设置默认值
else
	rs_total = rs.RecordCount	'统计记录总数
end if
dim pageno						'定义变量
getpageno = Switch(Request("pageno"))	'为getpageno变量赋值
if(getpageno = "")then					'判断getpageno变量是否为空
	pageno = 1							'设置默认值
else
	pageno = getpageno					'为pageno变量赋值
End if
if(not rs.Eof)then
	rs.AbsolutePage = pageno
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<style type="text/css">
<!--
.STYLE3 {font-size: 9pt; color: #000000; }
-->
</style>
</head>
<style type="text/css">
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
</style>
<script language="javascript">
function goto(id)
{
	location.href = "Detail1.asp?id=" + id;
}
</script>
<!--#include file="top.asp"-->
<body>
<table width="800" height="529" border="0" align="center" cellpadding="0" cellspacing="0" background="images/new14.gif">
  <tr>
    <td height="86" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td width="46" height="289">&nbsp;</td>
    <td width="713" valign="top"><table width="693" height="70" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="693" height="70"><table width="662" border="1" align="center"  cellpadding="0" cellspacing="0" bordercolorlight="#F7AB58" bordercolordark="#ffffff">
          <tr>
            <td height="24" align="center" nowrap="nowrap"> 新闻主题</td>
            <td height="25" align="center" nowrap="nowrap"> 新闻内容</td>
            <td align="center" nowrap="nowrap"> 发表时间</td>
          </tr>
          <%
	repeat_rows = 0										'为repeat_rows变量赋值
	While(repeat_rows < rs.PageSize and Not rs.Eof)		
	%>
          <tr style="cursor:hand" onmouseover="this.style.background='#dddddd'" onmouseout="this.style.background='#ffffff'" onclick="goto(<%=rs("Id")%>)">
            <td width="230" height="25" align="center" nowrap="nowrap"><%=Server.HtmlenCode(rs("title"))%> </td>
            <td width="288" align="center" nowrap="nowrap"><%
		If(Len(rs("content")) > 50)Then											'判断信息内容的长度是否大于50
		Response.Write(Server.HTMLEncode(Mid(rs("content"),1,50))&"....")		'输出信息内容并带超链接
		Else
		Response.Write(Server.HTMLEncode(rs("content")))						'输出信息内容并带超链接
		End If
		%>            </td>
            <td width="167" align="center" nowrap="nowrap"><%=rs("time1")%></td>
          </tr>
          <%
	repeat_rows = repeat_rows + 1												'将repeat_rows变量加1
	rs.MoveNext																	'向下移动记录指针
	Wend
	%>
        </table>
          <table cellspacing="0" cellpadding="0" border="0" height="18" width="662" align="center">
            <tr>
              <td align="left" nowrap="nowrap"> [<%=pageno%>/<%=rs.PageCount%>]&nbsp;每页<%=rs.PageSize%>条&nbsp;共<%=rs_total%>条记录 </td>
              <td align="right" nowrap="nowrap"><%
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
					%>              </td>
            </tr>
          </table></td>
      </tr>
    </table></td>
    <td width="41">&nbsp;</td>
  </tr>
  <tr>
    <td height="13" colspan="3"><table width="100%" height="21" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="40" height="15">&nbsp;</td>
        <td width="19">&nbsp;</td>
        <td width="665">&nbsp;</td>
        <td width="40">&nbsp;</td>
        <td width="34">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="133" colspan="3"><table width="800" height="133" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="133">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
