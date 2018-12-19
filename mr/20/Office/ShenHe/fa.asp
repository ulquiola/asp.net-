<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file=../conn/conn.asp-->
<%
	if request("post")="true" then 
	call wri
	end if 
	function wri
		if request("title")<>"" and request("content")<>"" then 
			set rs=server.CreateObject("adodb.recordset")
			sql="select * from Tab_shehe"
			rs.open sql,conn,0,3
			rs.AddNew
			rs("title")=request("title")
			rs("content")=request("content")
			rs("name")=session("admin_name")
			rs("shen")=0
			rs("time1")=date()
			rs.update
			rs.close
			response.Redirect("chenggong.htm")
		else
			response.Write("<script language=javascript>alert('请把信息填写完整')</script>")
		end if 
	end function
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<style type="text/css">
<!--
@import url("table.css");
@import url("biaodan.css");
body,td,th {
	font-size: 12px;
}
body {
	background-color: #ffffff;
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
.style4 {
	font-size: 16px;
	font-weight: bold;
	color: #FFFFFF;
}
.style5 {
	font-size: 14px;
	color: #000000;
}
-->
</style>
</head>

<body>


<table width="814" height="505" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="488" valign="top" background="../Images/main.gif"><table width="71%" height="84%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td>
		
		
		<form name="form1" method="post" action="">
  <table width="695" height="415" border="0" align="center" cellpadding="0" cellspacing="0"  >
    
    <tr align="left"  >
      <td height="77" colspan="2">
        <table width="543" height="37" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td height="14" colspan="2">&nbsp;</td>
          </tr>
          <tr>
            <td width="25" height="22" valign="bottom"><img src="../Images/add.gif" width="20" height="18"></td>
            <td width="518" valign="bottom"><img src="../Images/f.gif" width="70" height="18"></td>
          </tr>
        </table>
        <span class="style4"><br>
        </span></td>
    </tr>
    <tr>
      <td width="57" height="33"><div align="right">标题：</div></td>
      <td width="618"> 
        <input name="title" type="text" class="unnamed1" id="title" size="50">
      <input name="post" type="hidden" id="post" value="true"></td>
    </tr>
    <tr>
      <td width="57" height="113" valign="top" ><div align="right">内容：</div></td>
      <td> 
      <textarea name="content" cols="76" rows="20" id="content" class="unnamed1"></textarea></td>
    </tr>
    <tr>
      <td height="35" colspan="2"><div align="center">
        <input type="submit" name="Submit" value="提交">
        　
        <input type="reset" name="Submit2" value="重置">
      </div></td>
    </tr>
  </table>
</form>

		</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
