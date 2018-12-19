<!--#include file="../common/conn.asp"-->
<!--#include file="../Replace.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RecordSet")
note_id = request("note_id")					'获取超级链接中参数
sql = "select * from tb_note where note_id='"&note_id&"' and  note_flag=1"
rs.Open sql,conn,1,3
%>
<HTML>
<HEAD>
<TITLE>版主浏览悄悄话</TITLE>
<meta http-equiv="Content-Type" content="text/html;	charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<style>
.bg_author
{
	background-color:#eeeeee;
	height:24px;
}
.generalwindow {	MARGIN-TOP: 20px; WIDTH: 100%; BORDER-COLLAPSE: collapse
}
.outerborder {	BORDER-RIGHT: #ff9cbf 1px solid; PADDING-RIGHT: 2px; BORDER-TOP: #ff9cbf 1px solid; PADDING-LEFT: 2px; BACKGROUND-IMAGE: none; PADDING-BOTTOM: 2px; MARGIN: 0px auto; BORDER-LEFT: #ff9cbf 1px solid; WIDTH: 778px; PADDING-TOP: 2px; BORDER-BOTTOM: #ff9cbf 1px solid
}
.STYLE14 {color: #FF0000}
</style>
</HEAD>
<BODY>
<table width="778px" align="center"	cellpadding="0" cellspacing="0" bgcolor="" style="border:0;margin:10;align:center;">
  <tr height='32px' class='toptitle1' bgcolor='#ccddee'>
    <td width='11' bgcolor="#D9EBB9"></td>
    <td height="27" colspan="2" bgcolor="#D9EBB9"><!--回复相对留言的缩进距离-->
      <strong><font color="#559380"><%=rs("note_user")%></font></strong>&nbsp;&nbsp;<span class="style1 STYLE14">给版主的悄悄话</span> </td>
    <td width="23" bgcolor="#D9EBB9">&nbsp;</td>
  </tr>
  <tr height='26px'>
    <!-- 标题行 -->
    <td bgcolor="#F9F8EF"></td>
    <td colspan="2" bgcolor="#F9F8EF"><span class="style1"><%=rs("note_title")%></span> </td>
    <td bgcolor="#F9F8EF"><span class="style2"> <a href="note_answer.asp?note_id=<% =rs("note_id") %>"></a></span></td>
  </tr>
  <tr height='26px'>
    <!-- 内容行 -->
    <td bgcolor="#F9F8EF"></td>
    <td width="683" valign="top" bgcolor="#F9F8EF"><%=rs("note_content")%></td>
    <td width="59" bgcolor="#F9F8EF" style="padding-bottom:10px"><img src="../images/face/pic/<%=rs("note_user_pic")%>"><br>
    </td>
    <td bgcolor="#F9F8EF">&nbsp;</td>
  </tr>
  <tr class="bg_author">
    <!-- 留言人信息 -->
    <td></td>
    <td colspan="2" align="right">留言时间：&nbsp;<%=rs("note_time")%></td>
    <td align="right">&nbsp;</td>
  </tr>
</table>
</BODY>
</HTML>
