<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="checklogin.asp"-->
<%
If Not Isempty(Request("delete")) Then
  id=Request.Form("id")
  Set rs=conn.Execute("select Ppic from tab_photo where id="&id&"")
  pic=rs("Ppic")
  Set rs=Nothing
  pic="/upfile/"&pic
  Set FSObject=Server.CreateObject("Scripting.FileSystemObject")	
  If FSObject.FileExists(Server.MapPath(pic)) Then
    Set FileObject=FSObject.GetFile(Server.MapPath(pic))
	FileObject.Delete True
  End If
  sqlstr="delete from tab_photo where id="&id&""
  conn.Execute(sqlstr)
  Response.Redirect("ad_photo_list.asp")
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
                <td height="22"><img src="images/dian.gif" width="7" height="7"> ���鿴 </td>
              </tr>
              <tr>
                <td height="1"><img src="images/xian.gif" width="366" height="1"></td>
              </tr>
            </table><table width="90%" border="0" align="center" cellpadding="2" cellspacing="0">
        <form name="form1" method="get" action="">
          <tr>
            <td height="22" align="right">���: </td>
            <td><select name="txt_class" id="select">
                <option selected>ѡ�����</option>
                <%
				Set rs=Server.CreateObject("ADODB.Recordset")
				sqlstr="select id,Pcname from tab_photo_class"
				rs.open sqlstr,conn,1,1
				while not rs.eof
				%>
                <option value="<%=rs("id")%>"><%=rs("Pcname")%></option>
                <%
				rs.movenext
				wend
				rs.close
				Set rs=Nothing
				%>
            </select></td>
            <td height="22" align="right">ͼƬ����: </td>
            <td><input name="txt_title" type="text" class="textbox" id="txt_title" size="15"></td>
            <td height="22"><input name="query" type="submit" class="button" id="query" value="�� ѯ"></td>
          </tr>
        </form>
      </table><table width="90%" border="0" align="center" cellpadding="0" cellspacing="2">
        <tr align="center">
          <td width="134">���</td>
          <td width="246" height="22">ͼƬ����</td>
          <td width="170" height="22">�� ��</td>
        </tr>
<%
txt_class=Request("txt_class")
txt_title=Request("txt_title")
Set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select * from tab_photo where 1=1"
If txt_class<>"" and txt_class<>"ѡ�����" Then sqlstr=sqlstr&" and Pclass="&txt_class&""
If txt_title<>"" Then sqlstr=sqlstr&" and Pname like '%"&txt_title&"%'"
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
        <form name="form1<%=rs("id")%>" method="post" action="">
          <tr align="center">
            <td align="left">
              <%Set rsc=conn.Execute("select Pcname from tab_photo_class where id="&rs("Pclass")&"")
	  If Not rsc.eof or Not rsc.bof Then
	    Response.Write(rsc("Pcname"))
	  End If
	  Set rsc=Nothing
	%></td>
            <td height="22" align="left"><%=Left(rs("Pname"),15)%></td>
            <td height="22"><input name="id" type="hidden" id="id" value="<%=rs("id")%>">
                <a href="ad_photo.asp?id=<%=rs("id")%>&action=view">�޸�</a>
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
            <td height="22" colspan="3">
              <%	
	if pages<>1 then
		response.Write("&nbsp;&nbsp;<a href="&path&"?pages=1&txt_class="&txt_class&"&txt_title="&txt_title&">��ҳ</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?pages="&(pages-1)&"&txt_class="&txt_class&"&txt_title="&txt_title&">��һҳ</a>")
	end if 
	response.Write("&nbsp;&nbsp;��ǰ <font color='#FF0000'>"&pages&"/"&rs.pagecount&"</font> ҳ")
	if pages<>rs.pagecount then
		response.Write("&nbsp;&nbsp;<a href="&path&"?pages="&(pages+1)&"&txt_class="&txt_class&"&txt_title="&txt_title&">��һҳ</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?pages="&rs.pagecount&"&txt_class="&txt_class&"&txt_title="&txt_title&">ĩҳ</a>")
	end if
    rs.close
    Set rs=Nothing
   %>
            </td>
          </form>
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
