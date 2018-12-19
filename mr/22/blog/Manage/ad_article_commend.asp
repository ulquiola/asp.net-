<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="checklogin.asp"-->
<%
If Not Isempty(Request("delete")) Then
  id=Request.Form("id")
  sqlstr="delete from tab_article_commend where id="&id&""
  conn.Execute(sqlstr) 
  Response.Redirect("ad_article_list.asp")
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>博客</title>
<link rel="stylesheet" href="css/css.css">
</head>
<body>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><!--#include file="top.asp"--></td>
  </tr>
  <tr valign="top">
    <td height="5" background="images/mid_beijing.jpg"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="29%" align="center" valign="top"><!--#include file="left.asp"--></td>
          <td width="71%" valign="top"><table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td height="22"><img src="images/dian.gif" width="7" height="7"> 文章相关评论</td>
              </tr>
              <tr>
                <td height="1"><img src="images/xian.gif" width="366" height="1"></td>
              </tr>
            </table><table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <%
id=Request.QueryString("id")
If Not IsNumeric(id) Then id=1
Atitle=Request.QueryString("Atitle")
%>
        <tr>
          <td height="20">文章标题:<%=Atitle%></td>
        </tr>
        <tr>
          <td><table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#0099FF">
              <%
Set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select * from tab_article_commend where Cid="&id&""
rs.open sqlstr,conn,1,1
If rs.eof Then
%>
              <tr bgcolor="#FFFFFF">
                <td height="20" align="center" colspan="3">此文章暂无评论！</td>
              </tr>
              <%End IF%>
<%
If Not (rs.eof and rs.bof) Then
	rs.pagesize=5  '定义每页显示的记录数
	pages=clng(Request("pages"))  '获得当前页数
	If pages<1 Then pages=1
	If pages>rs.recordcount Then pages=rs.recordcount
	showpage rs,pages  '执行分页子程序showpage		
	Sub showpage(rs,pages)  '分页子程序showpage(rs,pages)
	rs.absolutepage=pages   '指定指针所在的当前位置
	For i=1 to rs.pagesize  '循环显示记录集中的记录
%>
              <form name="form<%=rs("id")%>" method="post" action="">
                <tr bgcolor="#FFFFFF">
                  <td height="22">评论人昵称:<%=rs("Cname")%></td>
                  <td height="22">评论时间:<%=rs("Cdate")%></td>
                  <td><input name="delete" type="submit" class="button" id="delete2" onClick="return confirm('确定要删除吗?')" value="删除">
                  <input name="id" type="hidden" id="id" value="<%=rs("id")%>"></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="22" colspan="3">评论内容:<%=rs("Ccontent")%></td>
                </tr>                
              </form>
<%
  rs.movenext  '指针向下移动
  If rs.eof Then exit for
  Next
  End Sub
End If
%>
		  <tr align="center" bgcolor="#FFFFFF">
		  <form name="form" action="?" method="get">
            <td height="22" colspan="3">
              <%	
				if pages<>1 then
					response.Write("&nbsp;&nbsp;<a href="&path&"?pages=1&id="&id&">首页</a>")
					response.Write("&nbsp;&nbsp;<a href="&path&"?pages="&(pages-1)&"&id="&id&">上一页</a>")
				end if 
				response.Write("&nbsp;&nbsp;当前 <font color='#FF0000'>"&pages&"/"&rs.pagecount&"</font> 页")
				if pages<>rs.pagecount then
					response.Write("&nbsp;&nbsp;<a href="&path&"?pages="&(pages+1)&"&id="&id&">下一页</a>")
					response.Write("&nbsp;&nbsp;<a href="&path&"?pages="&rs.pagecount&"&id="&id&">末页</a>")
				end if
				rs.close
				Set rs=Nothing
			  %></td>
          </form>
          </tr>
          </table></td>
        </tr>
      </table>
          </td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><!--#include file="bottom.asp"--></td>
  </tr>
</table>
</body>
</html>
