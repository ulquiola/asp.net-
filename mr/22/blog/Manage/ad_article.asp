<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="checklogin.asp"-->
<%
'����¼�¼
Sub add()
  str1=Str_filter(Request.Form("�������")) 
  str2=Str_filter(Request.Form("��������"))
  str3=Str_filter(Request.Form("��������"))
  str4=Str_filter(Request.Form("��������"))
If str1<>"" and str2<>"" and str3<>"" and str4<>"" Then 
  Set rs=Server.CreateObject("ADODB.Recordset")
  sqlstr="select * from tab_article"
  rs.open sqlstr,conn,1,3
  rs.addnew
  rs("Aclass")=str1
  rs("Aauthor")=str2
  rs("Atitle")=str3
  rs("Acontent")=str4
  rs.update
  rs.close
  Set rs=Nothing
  Response.Redirect("ad_article_list.asp")
Else
  Response.Write("<script>alert('����д����Ϣ������!');history.back();</script>")
End If
End Sub
'�޸ļ�¼
Sub edit()
  id=Request.Form("id")
  str1=Str_filter(Request.Form("�������")) 
  str2=Str_filter(Request.Form("��������"))
  str3=Str_filter(Request.Form("��������"))
  str4=Str_filter(Request.Form("��������"))
If str1<>"" and str2<>"" and str3<>"" and str4<>"" Then 
  Set rs=Server.CreateObject("ADODB.Recordset")
  sqlstr="select * from tab_article where id="&id&""
  rs.open sqlstr,conn,1,3
  rs("Aclass")=str1
  rs("Aauthor")=str2
  rs("Atitle")=str3
  rs("Acontent")=str4
  rs.update
  rs.close
  Set rs=Nothing
  Response.Redirect("ad_article_list.asp")
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
            <td height="22"><img src="images/dian.gif" width="7" height="7"> �������</td>
          </tr>
          <tr>
            <td height="1"><img src="images/xian.gif" width="366" height="1"></td>
          </tr>
        </table>
		<br>
    <%If Request.QueryString("action")="" Then%>
      <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
        <form name="form1" method="post" action="">
          <tr>
            <td height="28" align="right">�������:</td>
              <td height="28"><select name="�������" id="select">
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
          </tr>
          <tr>
            <td height="28" align="right">��������:</td>
              <td height="28"><input name="��������" type="text" class="textbox" id="��������"></td>
          </tr>
          <tr>
            <td height="28" align="right">��������:</td>
              <td height="28"><input name="��������" type="text" id="��������" class="textbox"></td>
          </tr>
          <tr>
            <td height="22" align="right">��������:</td>
              <td height="22"><textarea name="��������" cols="45" rows="6" id="��������"></textarea></td>
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
sqlstr="select * from tab_article where id="&id&""
Set rs=conn.Execute(sqlstr)
%>
        <form name="form2" method="post" action="">
          <tr>
            <td height="28" align="right">�������:</td>
              <td height="28"><select name="�������" id="�������">
                <option selected>ѡ�����</option>
                <%
				Set rsc=Server.CreateObject("ADODB.Recordset")
				sqlstr="select id,Acname from tab_article_class"
				rsc.open sqlstr,conn,1,1
				while not rsc.eof
				%>
                <option value="<%=rsc("id")%>" <%if rsc("id")=cint(rs("Aclass")) then Response.Write("selected") end if%>><%=rsc("Acname")%></option>
                <%
				rsc.movenext
				wend
				rsc.close
				Set rsc=Nothing
				%>
                </select></td>
          </tr>
          <tr>
            <td height="28" align="right">��������:</td>
              <td height="28"><input name="��������" type="text" class="textbox" id="��������" value="<%=rs("Aauthor")%>"></td>
          </tr>
          <tr>
            <td height="28" align="right">��������:</td>
              <td height="28"><input name="��������" type="text" id="��������" class="textbox" value="<%=rs("Atitle")%>"></td>
          </tr>
          <tr>
            <td height="22" align="right">��������:</td>
              <td height="22"><textarea name="��������" cols="45" rows="6" id="��������"><%=rs("Acontent")%></textarea></td>
          </tr>
          <tr>
            <td height="28" colspan="2" align="center">
              <input name="id" type="hidden" id="id" value="<%=rs("id")%>">
              <input name="edit" type="submit" class="button" id="edit" value="�� ��" onClick="return Mycheck(this.form)">
  &nbsp;
              <input type="button" name="Submit22" value="�� ��" class="button" onClick="javascript:window.location.href='ad_article_list.asp'">
            </td>
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
