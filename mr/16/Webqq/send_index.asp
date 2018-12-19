<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet">
<style type="text/css">
<!--
.STYLE2 {
	font-size: 9pt;
	color: #FF0000;
}
.STYLE4 {
	font-size: 9pt;
	color: #000000;
}
-->
</style>
</head>
<script language="javascript">
function mycheck(){
if(form2.Img_file.value=="")
{alert("请选择所要发送的图片！");form2.Img_file.focus();return false}
form2.submit();
}
</script>
<body>
<table width="607" height="37"  border="1" cellpadding="0" cellspacing="0" bordercolordark="#000000" bordercolorlight="#FFFFFF" bordercolor="#000000">
<tr>
      <td width="100%" height="35">
        <table>
	  <tr>
 <% 
	if request.QueryString("id")<>"" then
	ID=request.QueryString("id")
  	end if
	set rs_1=server.CreateObject("adodb.recordset")
	sql1="select * from tb_user where ID="&id
  	rs_1.open sql1,conn,1,3
%>	 
<form action="Upload_OK.asp?ID=<%=request.QueryString("ID")%>"" method="post" enctype="multipart/form-data" name="form2" target="mainFrame">
	   <td><span class="STYLE4">请选择图片：	  </span>
	  <input name="Img_file" type="file" id="Img_file" size="40">	  </td>

	  <td>
        <input type="button" name="SendImg" value="发送图片" onClick="mycheck();">
        <span class="STYLE4">给</span>        <span class="STYLE2">[<%=rs_1("username")%>]	    </span>
	  </td>
	  </form>
	  </tr>
	  </table>	</td>
  </tr>
</table>
</body>
</html>
