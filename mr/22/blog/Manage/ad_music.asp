<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="checklogin.asp"-->
<%
'����¼�¼
Sub add()
  str1=Str_filter(Request.Form("��������")) 
  str2=Str_filter(Request.Form("��������"))
  str3=Session("pic")
  str4=Str_filter(Request.Form("���"))
If str1<>"" and str2<>"" and str3<>"" and str4<>"" Then 
  Set rs=Server.CreateObject("ADODB.Recordset")
  sqlstr="select * from tab_music"
  rs.open sqlstr,conn,1,3
  rs.addnew
  rs("Mtitle")=str1
  rs("Mname")=str2
  rs("Mpath")=str3
  rs("Mtype")=UCase(right(Session("pic"),3))
  Session("pic")=""
  rs("Msize")=Session("upfile_size")
  Session("upfile_size")=""
  rs("Mword")=str4
  rs.update
  rs.close
  Set rs=Nothing
  Response.Redirect("ad_music_list.asp")
Else
  Response.Write("<script>alert('����д����Ϣ������!');history.back();</script>")
End If
End Sub


'�޸ļ�¼
Sub edit()
  id=Request.Form("id")
  str1=Str_filter(Request.Form("��������")) 
  str2=Str_filter(Request.Form("��������"))
  str3=Session("pic")
  str4=Str_filter(Request.Form("���"))  
If str1<>"" and str2<>"" and str4<>"" Then 
  Set rs=Server.CreateObject("ADODB.Recordset")
  sqlstr="select * from tab_music where id="&id&""
  rs.open sqlstr,conn,1,3
  rs("Mtitle")=str1
  rs("Mname")=str2
  If str3<>"" Then 
  path="/music/"&rs("Mpath")
  Set FSObject=Server.CreateObject("Scripting.FileSystemObject")	
  If FSObject.FileExists(Server.MapPath(path)) Then
    Set FileObject=FSObject.GetFile(Server.MapPath(path))
	FileObject.Delete True
  End If
  	rs("Mpath")=str3
	rs("Mtype")=UCase(right(Session("pic"),3))
    Session("pic")=""
	rs("Msize")=Session("upfile_size")
    Session("upfile_size")=""
  End If
  rs("Mword")=str4
  rs.update
  rs.close
  Set rs=Nothing
  Response.Redirect("ad_music_list.asp")
Else
  Response.Write("<script>alert('����д����Ϣ������!');history.back();</script>")
End If
End Sub
If Not Isempty(Request("add")) Then call add()
If Not Isempty(Request("edit")) Then call edit()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����</title>
<link rel="stylesheet" href="css/css.css">
<script language="javascript">
function Mycheck(form){
  for(i=0;i<form.length;i++){
    if(form.elements[i].value==""){
	  alert(form.elements[i].name + "����Ϊ��!");return false;}	
  }
}
</script>
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
                <td height="22"><img src="images/dian.gif" width="7" height="7"> ��Ƶ�ļ����</td>
              </tr>
              <tr>
                <td height="1"><img src="images/xian.gif" width="366" height="1"></td>
              </tr>
            </table>
	<%If Request.QueryString("action")="" Then%>
      <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
        <form name="form1" method="post" action="">
          <tr>
            <td width="82" height="28" align="right">��������:</td>
            <td width="418" height="28"><input name="��������" type="text" id="��������" class="textbox"></td>
          </tr>
          <tr>
            <td height="28" align="right">��������:</td>
            <td height="28"><input name="��������" type="text" id="��������" class="textbox"></td>
          </tr>
          <tr>
            <td height="22" align="right">�ļ�·��:</td>
            <td height="22"><div align="left">
                <iframe src="UpFile.asp" width="300" height="22" scrolling="no" MARGINHEIGHT="0" MARGINWIDTH="0" align="middle" frameborder="0"></iframe><BR>
                (�ϴ��ļ���ʽΪmp3 ��wav��mid����rmi,�ϴ��ļ���С��Ҫ����50K)
            </div></td>
          </tr>
          <tr>
            <td height="22" align="right">���:</td>
            <td height="22"><textarea name="���" cols="40" rows="4" id="���"></textarea></td>
          </tr>
          <tr>
            <td height="28" colspan="2" align="center"><input name="add" type="submit" class="button" id="add" value="�� ��" onClick="return Mycheck(this.form)">
&nbsp;
        <input type="reset" name="Submit2" value="�� ��" class="button"></td>
          </tr>
        </form>
      </table>
      <%Else%>
      <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
		<%
		id=Request.QueryString("id")
		sqlstr="select * from tab_music where id="&id&""
		Set rs=conn.Execute(sqlstr)
		%>
        <form name="form2" method="post" action="">
          <tr>
            <td width="87" height="22" align="right">��������:</td>
            <td width="413" height="22"><input name="��������" type="text" id="��������" class="textbox" value="<%=rs("Mtitle")%>"></td>
          </tr>
          <tr valign="top">
            <td height="22" align="right" valign="middle">��������:</td>
            <td height="22"><input name="��������" type="text" id="��������" class="textbox" value="<%=rs("Mname")%>"></td>
          </tr>
          <tr>
            <td height="22" align="right">�ļ�·��:</td>
            <td height="22"><div align="left">
                <iframe src="UpFile.asp" width="300" height="22" scrolling="no" MARGINHEIGHT="0" MARGINWIDTH="0" align="middle" frameborder="0"></iframe><BR>
                (�ϴ��ļ���ʽΪmp3 ��wav��mid����rmi,�ϴ��ļ���С��Ҫ����50K) </div></td>
          </tr>
          <tr>
            <td align="right">���:</td>
            <td><textarea name="���" cols="40" rows="4" id="���"><%=rs("Mword")%></textarea></td>
          </tr>
          <tr>
            <td height="28" align="center">&nbsp;</td>
            <td height="28" align="center"><input name="id" type="hidden" id="id" value="<%=rs("id")%>">
              <input name="edit" type="submit" class="button" id="edit" value="�� ��" onClick="return Mycheck(this.form)">
&nbsp;
<input type="button" name="Submit22" value="�� ��" class="button" onClick="javascript:window.location.href='ad_music_list.asp'"></td>
          </tr>
        </form>
      </table>
   		 <%End If%>
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
