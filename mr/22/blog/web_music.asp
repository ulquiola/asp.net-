<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<%
If Not Isempty(Request("play")) Then
  Dim choose,path,mydb,rs,sql,fs,mp3
  checkid=request.Form("checkid")
  If checkid="" then
    Response.Redirect("web_music.asp")
    Response.End()
  End if
  path=server.MapPath("music")
  set fs=server.CreateObject("scripting.filesystemobject")
  set mp3=fs.createtextfile(path&"\list.m3u",true)
  set rs=server.CreateObject("adodb.recordset")
  SQL="select Mpath,Mnum from tab_music where id in ("&checkid&")"
  rs.open sql,conn,1,3
  do while not rs.eof '���û�ѡ��ĸ����б�д��.m3u�ļ���
    rs("Mnum")=rs("Mnum")+1
    path=left(request.ServerVariables("URL"),instrrev(request.ServerVariables("URL"),"/"))
	If Request.ServerVariables("SERVER_PORT")<>"" Then
	  url_str="http://"&request.ServerVariables("LOCAL_ADDR")&":"&Request.ServerVariables("SERVER_PORT")&"/music/"&rs("Mpath")&chr(10)&chr(13)
	Else
	  url_str="http://"&request.ServerVariables("LOCAL_ADDR")&"/music/"&rs("Mpath")&chr(10)&chr(13)
	End If
    mp3.write(url_str)
  rs.update
  rs.movenext
  loop
  set rs=nothing
  mp3.close
  set mp3=nothing
  Response.redirect("\music\list.m3u")
  Response.end()
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css" />
<title>����</title>
</head>
<body>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><!--#include file="top.asp"--></td>
  </tr>
  <tr>
    <td bgcolor="FFFEED">
	<table width="70%" height="400"  border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="FFCC33">
      <tr>
        <td valign="top" bgcolor="#FFFFFF"><table width="700"  border="0" align="center" cellpadding="0" cellspacing="2">
          <form name="form1" method="post" action="">
            <tr>
              <td width="147" height="22" align="center"><input name="s_type" type="radio" value="0" checked>
        ��������</td>
              <td width="147" height="22" align="center"><input type="radio" name="s_type" value="1">
        ��������</td>
              <td width="220" height="22"><input name="keyword" type="text" class="textbox" id="keyword"></td>
              <td width="186" height="22"><input name="Submit" type="submit" class="button" value="�� ��"></td>
            </tr>
          </form>
        </table>
          <table width="700" border="0" align="center" cellpadding="0" cellspacing="2">
            <tr align="center" bgcolor="FFFCE8">
              <td height="22">��������</td>
              <td height="22">��������</td>
              <td height="22">��С</td>
              <td height="22">��ʽ</td>
              <td height="22">���</td>
              <td height="22">�����</td>
              <td height="22">�㲥</td>
            </tr>
            <form name="form2" method="post" action="">
              <%
s_type=Request("s_type")
keyword=Request("keyword")
Set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select * from tab_music where 1=1"
Select case s_type
case "0"
  sqlstr=sqlstr&" and Mtitle like '%"&keyword&"%'"
case "1"
  sqlstr=sqlstr&" and Mname like '%"&keyword&"%'"
End Select
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
              <tr align="left">
                <td height="20" align="center"><a href="web_music_download.asp?id=<%=rs("id")%>"><%=rs("Mtitle")%></a></td>
                <td height="20" align="center"><%=rs("Mname")%></td>
                <td height="20" align="center"><%=rs("Msize")%></td>
                <td height="20" align="center"><%=rs("Mtype")%></td>
                <td height="20" align="center"><a href="web_music_word.asp?id=<%=rs("id")%>" target="_blank">���</a></td>
                <td height="20" align="center"><%=rs("Mnum")%></td>
                <td height="20" align="center"><input name="checkid" type="checkbox" id="checkid3" value="<%=rs("id")%>"></td>
              </tr>
              <%
  rs.movenext  'ָ�������ƶ�
  If rs.eof Then exit for
  Next
  End Sub
End If
%>
              <tr align="center">
                <td height="22" colspan="7"><input name="play" type="submit" class="submit" id="play3" value="��ʼ����"></td>
              </tr>
            </form>
            <tr>
                <td height="22" colspan="7" align="center"></td>
            </tr>
          </table></td>
      </tr>
      <tr>
        <td height="22" bgcolor="FFFFCC"><table width="700" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="390" height="25" align="right"></td>
                      <td width="310" align="center"><%	
	if pages<>1 then
		response.Write("&nbsp;&nbsp;<a href="&path&"?pages=1&s_type="&s_type&"&keyword="&keyword&">��ҳ</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?pages="&(pages-1)&"&s_type="&s_type&"&keyword="&keyword&">��һҳ</a>")
	end if 
	response.Write("&nbsp;&nbsp;��ǰ <font color='#FF0000'>"&pages&"/"&rs.pagecount&"</font> ҳ")
	if pages<>rs.pagecount then
		response.Write("&nbsp;&nbsp;<a href="&path&"?pages="&(pages+1)&"&s_type="&s_type&"&keyword="&keyword&">��һҳ</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?pages="&rs.pagecount&"&s_type="&s_type&"&keyword="&keyword&">ĩҳ</a>")
	end if
    rs.close
    Set rs=Nothing
   %></td>
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
