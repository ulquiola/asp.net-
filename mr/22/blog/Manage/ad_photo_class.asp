<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="checklogin.asp"-->
<%
'添加新记录
Sub add()
  str1=Str_filter(Request.Form("类别名称"))  
  sqlstr="insert into tab_photo_class(Pcname) values('"&str1&"')"  
  conn.Execute(sqlstr)
  Response.Redirect("ad_photo_class.asp")
End Sub
'修改记录
Sub edit()
  str1=Str_filter(Request.Form("类别名称")) 
  id=Request.Form("id")  
  sqlstr="update tab_photo_class set Pcname='"&str1&"' where id="&id&""
  conn.Execute(sqlstr)  
  Response.Redirect("ad_photo_class.asp")
End Sub
'删除记录
Sub del()
  id=Request.Form("id")
  conn.Execute("delete from tab_photo where Pclass="&id&"")
  sqlstr="delete from tab_photo_class where id="&id&""
  conn.Execute(sqlstr)
  Response.Redirect("ad_photo_class.asp")
End Sub
'执行子过程
If Not Isempty(Request("add")) Then call add()
If Not Isempty(Request("edit")) Then call edit()
If Not Isempty(Request("delete")) Then call del()
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
                <td height="22"><img src="images/dian.gif" width="7" height="7"> 相册分类 </td>
              </tr>
              <tr>
                <td height="1"><img src="images/xian.gif" width="366" height="1"></td>
              </tr>
            </table><table width="500" border="0" align="center" cellpadding="0" cellspacing="2">
        <tr align="center">
          <td height="22">类别名称</td>
          <td height="22">操 作</td>
        </tr>
        <%
Set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select id,Pcname from tab_photo_class"
rs.open sqlstr,conn,1,1
while not rs.eof
%>
        <form name="form2<%=rs("id")%>" method="post" action="">
          <tr align="center">
            <td height="22"><input name="类别名称" type="text" id="类别名称2" value="<%=rs("Pcname")%>" class="textbox"></td>
            <td height="22"><input name="id" type="hidden" id="id" value="<%=rs("id")%>">
                <input name="edit" type="submit" id="edit" value="修 改" class="button" onClick="return Mycheck(this.form)">
&nbsp;
        <input name="delete" type="submit" id="delete" value="删 除" onClick="return confirm('确定要删除吗?')" class="button"></td>
          </tr>
        </form>
        <%
rs.movenext
wend
rs.close
set rs=nothing
%>
      </table><table width="500" border="0" align="center" cellpadding="2" cellspacing="0">
        <form name="form1" method="post" action="">
          <tr>
            <td width="106" height="22" align="right">类别名称:</td>
            <td width="253" height="22"><input name="类别名称" type="text" id="类别名称" class="textbox"></td>
            <td width="141" height="22"><input name="add" type="submit" id="add" value="添 加" class="button" onClick="return Mycheck(this.form)"></td>
          </tr>
        </form>
    </table>
          </td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><!--#include file="bottom.asp"--></td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
</html>
