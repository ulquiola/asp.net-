<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="checklogin.asp"-->
<%
If Not Isempty(Request("delete")) Then
  id=Request.Form("id")
  sqlstr="delete from tab_article_commend where Cid="&id&""
  conn.Execute(sqlstr)
  sqlstr="delete from tab_article where id="&id&""
  conn.Execute(sqlstr)
  Response.Redirect("ad_article_list.asp")
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����</title>
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
        <td height="22"><img src="images/dian.gif" width="7" height="7"> �����б�</td>
      </tr>
      <tr>
        <td height="1"><img src="images/xian.gif" width="366" height="1"></td>
      </tr>
    </table>
    <table width="550" border="0" align="center" cellpadding="0" cellspacing="2">
        <form name="form1" method="get" action="">
          <tr>
            <td height="22" align="right">���: </td>
            <td><select name="txt_class" id="select">
                <option selected>ѡ�����</option>
                <%
				Set rs=Server.CreateObject("ADODB.Recordset")
				sqlstr="select id,Acname from tab_article_class"
				rs.open sqlstr,conn,1,1
				while not rs.eof
				%>
                <option value="<%=rs("id")%>"><%=rs("Acname")%></option>
                <%
				rs.movenext
				wend
				rs.close
				Set rs=Nothing
				%>
            </select></td>
            <td height="22" align="right">���±���: </td>
            <td><input name="txt_title" type="text" class="textbox" id="txt_title" size="15"></td>
            <td height="22">&nbsp;</td>
          </tr>
          <tr>
            <td height="22" align="right">��������:</td>
            <td>
              <input name="txt_author" type="text" class="textbox" id="txt_author" size="12">
            </td>
            <td height="22" align="right">&nbsp;</td>
            <td>&nbsp;</td>
            <td height="22"><input name="query" type="submit" class="button" id="query" value="�� ѯ"></td>
          </tr>
        </form>
      </table>
      <table width="550" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr align="center">
          <td width="86">���</td>
          <td width="77" height="22">����</td>
          <td width="191" height="22">���±���</td>
          <td width="146" height="22">�� ��</td>
        </tr>
<%
txt_class=Request("txt_class")
txt_title=Request("txt_title")
txt_author=Request("txt_author")
Set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select * from tab_article where 1=1"
If txt_class<>"" and txt_class<>"ѡ�����" Then sqlstr=sqlstr&" and Aclass="&txt_class&""
If txt_title<>"" Then sqlstr=sqlstr&" and Atitle like '%"&txt_title&"%'"
If txt_author<>"" Then sqlstr=sqlstr&" and Aauthor like '%"&txt_author&"%'"
sqlstr=sqlstr&" order by id desc"
rs.open sqlstr,conn,1,1
If Not (rs.eof and rs.bof) Then
	rs.pagesize=8  '����ÿҳ��ʾ�ļ�¼��
	pages=clng(Request("pages"))  '��õ�ǰҳ��
	If pages<1 Then pages=1
	If pages>rs.recordcount Then pages=rs.recordcount
	showpage rs,pages  'ִ�з�ҳ�ӳ���showpage		
	Sub showpage(rs,pages)  '��ҳ�ӳ���showpage(rs,pages)
	rs.absolutepage=pages   'ָ��ָ�����ڵĵ�ǰλ��
	For i=1 to rs.pagesize  'ѭ����ʾ��¼���еļ�¼
%>
        <form name="form1" method="post" action="">
          <tr align="center">
            <td align="center">
            <%Set rsc=conn.Execute("select Acname from tab_article_class where id="&rs("Aclass")&"")
	  Response.Write(rsc("Acname"))
	  Set rsc=Nothing
	%></td>
            <td height="22" align="center"><%=rs("Aauthor")%></td>
            <td height="22" align="left"><%=Left(rs("Atitle"),15)%></td>
            <td height="22"><input name="id" type="hidden" id="id" value="<%=rs("id")%>">
                <a href="ad_article.asp?id=<%=rs("id")%>&action=view">�޸�</a> <a href="ad_article_commend.asp?id=<%=rs("id")%>&Atitle=<%=rs("Atitle")%>">����</a>
                <input name="delete" type="submit" id="delete" value="ɾ��" onClick="return confirm('ȷ��Ҫɾ����?')" class="button"></td>
          </tr>
        </form>
<%
  rs.movenext  'ָ�������ƶ�
  If rs.eof Then exit for
  Next
  End Sub
End If
%>
        <tr align="center">
          <form name="form" action="?" method="get">
            <td height="22" colspan="4">
              <%	
	if pages<>1 then
		response.Write("&nbsp;&nbsp;<a href="&path&"?pages=1&txt_class="&txt_class&"&txt_title="&txt_title&"&txt_author="&txt_author&">��ҳ</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?pages="&(pages-1)&"&txt_class="&txt_class&"&txt_title="&txt_title&"&txt_author="&txt_author&">��һҳ</a>")
	end if 
	response.Write("&nbsp;&nbsp;��ǰ <font color='#FF0000'>"&pages&"/"&rs.pagecount&"</font> ҳ")
	if pages<>rs.pagecount then
		response.Write("&nbsp;&nbsp;<a href="&path&"?pages="&(pages+1)&"&txt_class="&txt_class&"&txt_title="&txt_title&"&txt_author="&txt_author&">��һҳ</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?pages="&rs.pagecount&"&txt_class="&txt_class&"&txt_title="&txt_title&"&txt_author="&txt_author&">ĩҳ</a>")
	end if
    rs.close
    Set rs=Nothing
   %>
            </td>
          </form>
        </tr>
      </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><!--#include file="bottom.asp"--></td>
  </tr>
</table>
</body>
</html>
