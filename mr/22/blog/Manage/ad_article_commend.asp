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
                <td height="22"><img src="images/dian.gif" width="7" height="7"> �����������</td>
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
          <td height="20">���±���:<%=Atitle%></td>
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
                <td height="20" align="center" colspan="3">�������������ۣ�</td>
              </tr>
              <%End IF%>
<%
If Not (rs.eof and rs.bof) Then
	rs.pagesize=5  '����ÿҳ��ʾ�ļ�¼��
	pages=clng(Request("pages"))  '��õ�ǰҳ��
	If pages<1 Then pages=1
	If pages>rs.recordcount Then pages=rs.recordcount
	showpage rs,pages  'ִ�з�ҳ�ӳ���showpage		
	Sub showpage(rs,pages)  '��ҳ�ӳ���showpage(rs,pages)
	rs.absolutepage=pages   'ָ��ָ�����ڵĵ�ǰλ��
	For i=1 to rs.pagesize  'ѭ����ʾ��¼���еļ�¼
%>
              <form name="form<%=rs("id")%>" method="post" action="">
                <tr bgcolor="#FFFFFF">
                  <td height="22">�������ǳ�:<%=rs("Cname")%></td>
                  <td height="22">����ʱ��:<%=rs("Cdate")%></td>
                  <td><input name="delete" type="submit" class="button" id="delete2" onClick="return confirm('ȷ��Ҫɾ����?')" value="ɾ��">
                  <input name="id" type="hidden" id="id" value="<%=rs("id")%>"></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="22" colspan="3">��������:<%=rs("Ccontent")%></td>
                </tr>                
              </form>
<%
  rs.movenext  'ָ�������ƶ�
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
					response.Write("&nbsp;&nbsp;<a href="&path&"?pages=1&id="&id&">��ҳ</a>")
					response.Write("&nbsp;&nbsp;<a href="&path&"?pages="&(pages-1)&"&id="&id&">��һҳ</a>")
				end if 
				response.Write("&nbsp;&nbsp;��ǰ <font color='#FF0000'>"&pages&"/"&rs.pagecount&"</font> ҳ")
				if pages<>rs.pagecount then
					response.Write("&nbsp;&nbsp;<a href="&path&"?pages="&(pages+1)&"&id="&id&">��һҳ</a>")
					response.Write("&nbsp;&nbsp;<a href="&path&"?pages="&rs.pagecount&"&id="&id&">ĩҳ</a>")
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
