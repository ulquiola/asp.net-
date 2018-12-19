<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
'获取日志数
sqlstr="select count(id) from tab_article"
Set rs=conn.Execute(sqlstr)
num1=rs.fields(0).value
Set rs=Nothing
'获取评论数
sqlstr=sql&"select count(id) from tab_article_commend"
Set rs=conn.Execute(sqlstr)
num2=rs.fields(0).value
Set rs=Nothing
'获取访问人数
Set FSObject=Server.CreateObject("Scripting.FileSystemObject")
Set TextFile=FSObject.OpenTextFile(Server.MapPath("count.txt"))
num3=TextFile.ReadLine
Set TextFile=Nothing
Set FSObject=Nothing
%>
<table width="200" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="89">&nbsp;&nbsp;日志:</td>
    <td width="111"><%=num1%> 篇</td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp;评论:</td>
    <td><%=num2%> 个</td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp;访问:</td>
    <td><%=num3%> 人</td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp;建站时间:</td>
    <td>2008-6-26</td>
  </tr>
</table>
