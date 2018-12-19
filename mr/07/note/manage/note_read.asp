<!--#include file="../common/conn.asp"-->
<!--#include file="../Replace.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RecordSet")
note_id = request("note_id")					'获取超级链接中参数
sql = "select note.*,noan.* from tb_note as note left join tb_note_answer as noan on note.note_id = noan.noan_note_id where note.note_id ="&note_id &"order by note.note_time desc,noan.noan_time desc"
rs.Open sql,conn,1,3
%>
<HTML>
<HEAD>
<TITLE> 留言浏览 </TITLE>
<meta http-equiv="Content-Type" content="text/html;	charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<style>
.bg_author
{
	background-color:#eeeeee;
	height:24px;
}
.generalwindow { MARGIN-TOP: 20px; WIDTH: 100%; BORDER-COLLAPSE: collapse }
.outerborder{ BORDER-RIGHT: #ff9cbf 1px solid; PADDING-RIGHT: 2px; BORDER-TOP: #ff9cbf 1px solid; PADDING-LEFT: 2px; BACKGROUND-IMAGE: none; PADDING-BOTTOM: 2px; MARGIN: 0px auto; BORDER-LEFT: #ff9cbf 1px solid; WIDTH: 778px; PADDING-TOP: 2px; BORDER-BOTTOM: #ff9cbf 1px solid }
</style>
</HEAD>
<BODY>
<table width="778px" align="center" cellpadding="0" cellspacing="0" bgcolor="" style="border:0;margin:10;align:center;">
	<tr height='32px' class='toptitle1' bgcolor='#ccddee'>
	<td width='11' bgcolor="#D9EBB9"></td>
	<td height="27" bgcolor="#D9EBB9">
	  <span class="style1"><% =rs("note_title")%></span></td>
	<td width="15" bgcolor="#D9EBB9">&nbsp;</td>
  </tr>

  <tr height='26px'><!-- 标题行 -->
	<td bgcolor="#F9F8EF"></td>
	<td bgcolor="#F9F8EF"><span class="style1">
	  <%
	  note_user=rs("note_user")
	  if (note_flag<>1) then
		  response.write(note_user)
	  %>
	  ：说
	  <% end if %>
	</span></td>
	<td bgcolor="#F9F8EF">
	<span class="style2">
	<a href="note_answer.asp?note_id=<% =rs("note_id") %>"></a></span></td>
  </tr>
  <tr height='26px'><!-- 内容行 -->
	<td bgcolor="#F9F8EF"></td>
	<td bgcolor="#F9F8EF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="93%" valign="top" style="padding-left:10px; padding-right:10px;line-height:18px">&nbsp;&nbsp;&nbsp;&nbsp;<% =rs("note_content")%></td>
        <td width="7%" align="right" valign="middle" style="padding-bottom:10px"><img src="../images/face/pic/<% =rs("note_user_pic") %>"></td>
      </tr>
    </table></td>
    <td width="15" bgcolor="#F9F8EF">&nbsp;</td>
  </tr>
  <tr class="bg_author"><!-- 留言人信息 -->
	<td></td>
	<td align="right">留言时间：&nbsp;<%=rs("note_time") %></td>
    <td width="15" align="right">&nbsp;</td>
  </tr>
		<!-- 回复信息 -->
		<%
			if(rs("noan_id")<>"") then	'表示有回复
		%>
		<tr height='26px'>
		<td height="26px" bgcolor="#F9F8EF"></td>
		<td rowspan="2" align="center" bgcolor="#F9F8EF"><br>
		  <TABLE width="700" align="center" cellPadding=2 cellSpacing=1 bgcolor="#D9D2B6" class=embedbox >
          <TBODY>
            <TR>
              <TD width="700" bgcolor="#FFFFFF" style="padding-top:6px; padding-left:10px; padding-bottom:2px; padding-right:10px;line-height:18px"><SPAN 
                  style="FONT-WEIGHT: bold; COLOR: #000000">&nbsp;版主回复：</SPAN> <SPAN style="COLOR: #000000">(<% =rs("noan_time")%>)</SPAN>&nbsp;
                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="answer_del.asp?id=<% =rs("noan_id") %>&note_id=<% =rs("note_id")%>">删除</a>
                  <hr color=#D9D2B6 size=1 style="width:700px; " >
              <IMG 
                  src="../images/face/pic/01.gif" width="90" height="90" class=face style="FLOAT: left; MARGIN: 2px 5px 5px 2px">&nbsp;&nbsp;<SPAN 
                  style="COLOR: #000000"><% =rs("noan_content")%></SPAN></TD>
            </TR>
          </TBODY>
        </TABLE>
	      <br>
	      <br></td>	
		<td width="15" height="26px" bgcolor="#F9F8EF">&nbsp;</td>
		</tr>
		<tr><!-- 版主回复信息 -->
		<td bgcolor="#F9F8EF"></td>
		<td bgcolor="#F9F8EF">&nbsp;</td>
		</tr>
	<%end if%>
</table>

<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width="778px" style="border:0;align:center;" cellpadding="0" cellspacing="0" bgcolor="">
      <tr height='32px' class='toptitle1' bgcolor='#ccddee'>
        <td width='10' height="27" bgcolor="#D9EBB9"></td>
        <td width="100" height="27" bgcolor="#D9EBB9"><!--回复相对留言的缩进距离-->
              <span class="style1">版主回复</span> </td>
        <td width="500" height="27" bgcolor="#D9EBB9">&nbsp;</td>
        <td width="60" height="27" bgcolor="#D9EBB9">&nbsp;</td>
      </tr>
      <tr height='26px'>
        <!-- 标题行 -->
        <td height="13" colspan="4" bgcolor="#F9F8EF"></td>
      </tr>
    </table>
        <table width="778px" style="border:0;margin:0;"cellpadding="0" cellspacing="0" bgcolor="#FFF6F8">
          <form name="note" action="" method="post"	onsubmit="return(check_form())">
            <tr>
              <td width="111" height="24" align="center" bgcolor="F9F8EF">回复内容：<br></td>
              <td width="661" height="24" bgcolor="F9F8EF"><textarea name="content" rows="8" cols="80"></textarea>
              </td>
            </tr>
            <tr>
              <td height="35" colspan="2" align="center" bgcolor="F9F8EF"><input name="submit" type="submit" class="btn1" value=" 回复 " size="8">
                &nbsp;&nbsp;
                <input name="button" type="reset" class="btn1" onClick="clear_form('note')" value=" 清除 "size="8">
              </td>
            </tr>
          </form>
      </table></td>
  </tr>
</table>
<!--#include file="user.asp"-->
<%
if(request("submit")<>"") then
	content = request.Form("content")
	datetime=now()
	Set rs1 = Server.CreateObject("ADODB.RecordSet")
	sql1 = "insert tb_note_answer (noan_note_id,noan_content,noan_time,noan_user_name) values("&note_id&",'"&content&"','"&datetime&"','"&default_user_name&"')"
	rs1.Open sql1,conn,1,3
	Set rs2 = Server.CreateObject("ADODB.RecordSet")
	sql2= "update tb_note set note_answer=1 where note_id='"&note_id&"'"
	rs2.Open sql2,conn,1,3
%>
<script language="javascript">alert("留言信息回复成功！");window.location.href="note_read.asp?note_id="+<% response.Write(note_id)%>;</script>
<%
end if
%>
</BODY>
</HTML>
