<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../Conn/conn.asp" -->
<%'查询公告信息
if request("ID")<>"" then
	session("ID")=request("ID")
end if
Set rs_bbs = Server.CreateObject("ADODB.Recordset")
sql_bbs="SELECT * FROM Tab_Placard WHERE ID= "&session("ID")&""
rs_bbs.open sql_bbs,conn,1,3
%>
<%
'删除公告信息
if request.Form("subject")<>"" then
	Del="Delete From Tab_Placard where ID='"&session("ID")&"'"
	conn.execute(Del)
	response.Redirect("bbs_del_OK.asp")'将页面重定向到信息删除成功页面
end if%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>删除公告信息！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
.style1 {color: #C60001}
.style2 {color: #669999}
.STYLE4 {
	font-size: 9pt;
	color: #FFFFFF;
}
.STYLE6 {
	font-size: 9pt;
	color: #000000;
}
-->
</style>
</head>
<body>
<table width="557" height="401" border="0" cellpadding="-2" cellspacing="-2" 
background="../../Images/detail0.gif">
  <tr>
    <td width="817" height="349" valign="top"><table width="100%" height="76"
	  border="0" cellpadding="-2" cellspacing="-2">
      <tr>
        <td height="41" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE4">删除公告信息</span></td>
      </tr>
      <tr>
        <td height="35">&nbsp;</td>
      </tr>
    </table>      
      <form ACTION="bbs_del.asp" METHOD="POST" name="form1">
        <table width="80%" height="224"  border="0" align="center" cellpadding="-2" cellspacing="-2">
          <tr>
            <td height="27" colspan="2" align="left"><span class="style1">&nbsp;<span class="STYLE6">发布人：</span></span>
			<span class="STYLE6"><%=(rs_bbs.Fields.Item("Person").Value)%></span></td>
            <td width="19%"><div align="center" class="STYLE6">发布日期：</div></td>
            <td width="30%"><span class="STYLE6"><%=(rs_bbs.Fields.Item("dDate").Value)%></span></td>
          </tr>
          <tr>
            <td width="14%" height="29"><div align="center"><span class="STYLE6">主题：</span></div></td>
            <td colspan="3"><input name="subject" type="text" class="Style_subject" 
			id="subject2" value="<%=(rs_bbs.Fields.Item("Subject").Value)%>"></td>
          </tr>
          <tr>
            <td height="129" valign="top"><div align="center" class="STYLE6">内容：</div></td>
            <td colspan="3"><textarea name="content" cols="53" rows="6" class="Style_content"
			 id="content"><%=(rs_bbs.Fields.Item("Content").Value)%></textarea></td>
          </tr>
          <tr>
            <td colspan="4"><div align="center">
                <input name="but_del" type="submit" class="Style_button_del" id="del"
				 value="确定删除">
                <input name="myclose" type="button" class="Style_button" id="myclose"
				 value="关闭" onClick="javascrip:window.close()">
			</div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>
