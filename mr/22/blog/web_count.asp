<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
'��ȡ��־��
sqlstr="select count(id) from tab_article"
Set rs=conn.Execute(sqlstr)
num1=rs.fields(0).value
Set rs=Nothing
'��ȡ������
sqlstr=sql&"select count(id) from tab_article_commend"
Set rs=conn.Execute(sqlstr)
num2=rs.fields(0).value
Set rs=Nothing
'��ȡ��������
Set FSObject=Server.CreateObject("Scripting.FileSystemObject")
Set TextFile=FSObject.OpenTextFile(Server.MapPath("count.txt"))
num3=TextFile.ReadLine
Set TextFile=Nothing
Set FSObject=Nothing
%>
<table width="200" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="89">&nbsp;&nbsp;��־:</td>
    <td width="111"><%=num1%> ƪ</td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp;����:</td>
    <td><%=num2%> ��</td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp;����:</td>
    <td><%=num3%> ��</td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp;��վʱ��:</td>
    <td>2008-6-26</td>
  </tr>
</table>
