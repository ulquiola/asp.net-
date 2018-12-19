<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="checklogin.asp"-->
<%
'添加新记录
Sub add()
  str1=Str_filter(Request.Form("相册类别")) 
  str2=Str_filter(Request.Form("图片名称"))
  str3=Session("pic")
If str1<>"" and str2<>"" and str3<>"" Then 
  Set rs=Server.CreateObject("ADODB.Recordset")
  sqlstr="select * from tab_photo"
  rs.open sqlstr,conn,1,3
  rs.addnew
  rs("Pclass")=str1
  rs("Pname")=str2
  rs("Ppic")=str3
  Session("pic")=""
  rs.update
  rs.close
  Set rs=Nothing
  Response.Redirect("ad_photo_list.asp")
Else
  Response.Write("<script>alert('您填写的信息不完整!');history.back();</script>")
End If
End Sub
'修改记录
Sub edit()
  id=Request.Form("id")
  str1=Str_filter(Request.Form("相册类别")) 
  str2=Str_filter(Request.Form("图片名称"))
  str3=Session("pic")
If str1<>"" and str2<>"" Then 
  Set rs=Server.CreateObject("ADODB.Recordset")
  sqlstr="select * from tab_photo where id="&id&""
  rs.open sqlstr,conn,1,3
  rs("Pclass")=str1
  rs("Pname")=str2
  If str3<>"" Then 
  pic="/upfile/"&rs("Ppic")
  Set FSObject=Server.CreateObject("Scripting.FileSystemObject")	
  If FSObject.FileExists(Server.MapPath(pic)) Then
    Set FileObject=FSObject.GetFile(Server.MapPath(pic))
	FileObject.Delete True
  End If
  rs("Ppic")=str3
  Session("pic")=""
  End If
  rs.update
  rs.close
  Set rs=Nothing
  Response.Redirect("ad_photo_list.asp")
Else
  Response.Write("<script>alert('您填写的信息不完整!');history.back();</script>")
End If
End Sub
If Not Isempty(Request("add")) Then call add()
If Not Isempty(Request("edit")) Then call edit()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>博客</title>
<link rel="stylesheet" href="css/css.css">
<script language="javascript">
function Mycheck(form){
  for(i=0;i<form.length;i++){
    if(form.elements[i].value==""){
	  alert(form.elements[i].name + "不能为空!");return false;}	
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
        <td height="22"><img src="images/dian.gif" width="7" height="7"> 相册上传</td>
      </tr>
      <tr>
        <td height="1"><img src="images/xian.gif" width="366" height="1"></td>
      </tr>
    </table>
    <%If Request.QueryString("action")="" Then%>
      <table width="500" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#E3E3E3">
        <form name="form1" method="post" action="">
          <tr>
            <td width="82" height="28" align="right" bgcolor="#FFFFFF">相册类别:</td>
            <td width="418" height="28" bgcolor="#FFFFFF"><select name="相册类别" id="select">
                <option selected>选择类别</option>
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
          </tr>
          <tr>
            <td height="28" align="right" bgcolor="#FFFFFF">图片名称:</td>
            <td height="28" bgcolor="#FFFFFF"><input name="图片名称" type="text" id="图片名称" class="textbox"></td>
          </tr>
          <tr>
            <td height="22" align="right" bgcolor="#FFFFFF">图片信息:</td>
            <td height="22" bgcolor="#FFFFFF"><div align="left">
                <iframe src="UpFile.asp" width="300" height="22" scrolling="no" MARGINHEIGHT="0" MARGINWIDTH="0" align="middle" frameborder="0"></iframe>
            </div></td>
          </tr>
          <tr>
            <td height="28" colspan="2" align="center" bgcolor="#FFFFFF"><input name="add" type="submit" class="button" id="add" value="添 加" onClick="return Mycheck(this.form)">
&nbsp;
        <input type="reset" name="Submit2" value="重 置" class="button"></td>
          </tr>
        </form>
      </table>
      <%Else%>
      <table width="550" border="0" align="center" cellpadding="0" cellspacing="0">
        <%
id=Request.QueryString("id")
sqlstr="select * from tab_photo where id="&id&""
Set rs=conn.Execute(sqlstr)
%>
        <form name="form2" method="post" action="">
          <tr>
            <td width="118" rowspan="3" align="center"><img src="/upfile/<%=rs("Ppic")%>" height="80" width="100"></td>
            <td width="67" height="28" align="right">相册类别:</td>
            <td width="365" height="28"><select name="相册类别" id="相册类别">
                <option selected>选择类别</option>
                <%sqlstr="select id,Pcname from tab_photo_class"
				Set rsc=Server.CreateObject("ADODB.Recordset")
				sqlstr="select id,Pcname from tab_photo_class"
				rsc.open sqlstr,conn,1,1
				while not rsc.eof
				%>
                <option value="<%=rsc("id")%>" <%if rsc("id")=cint(rs("Pclass")) then Response.Write("selected") end if%>><%=rsc("Pcname")%></option>
                <%
				rsc.movenext
				wend
				rsc.close
				Set rsc=Nothing
				%>
            </select></td>
          </tr>
          <tr valign="top">
            <td height="28" align="right">图片名称:</td>
            <td height="28"><input name="图片名称" type="text" id="图片名称" class="textbox" value="<%=rs("Pname")%>">            </td>
          </tr>
          <tr>
            <td height="22" align="right">图片信息:</td>
            <td height="22"><div align="left">
                <iframe src="UpFile.asp" width="300" height="22" scrolling="no" MARGINHEIGHT="0" MARGINWIDTH="0" align="middle" frameborder="0"></iframe>
            </div></td>
          </tr>
          <tr>
            <td height="28" colspan="3" align="center"><input name="id" type="hidden" id="id" value="<%=rs("id")%>">
                <input name="edit" type="submit" class="button" id="edit" value="修 改" onClick="return Mycheck(this.form)">
&nbsp;
        <input type="button" name="Submit22" value="返 回" class="button" onClick="javascript:window.location.href='ad_photo_list.asp'"></td>
          </tr>
        </form>
      </table>
    <%End If%></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><!--#include file="bottom.asp"--></td>
  </tr>
</table>
</body>
</html>
