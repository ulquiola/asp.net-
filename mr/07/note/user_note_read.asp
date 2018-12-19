<!--#include file="common/conn.asp"-->
<%
	Set rs = Server.CreateObject("ADODB.RecordSet")
	note_id = request("id")
	sql = "select note.*,noan.* from tb_note as note left join tb_note_answer as noan on note.note_id = noan.noan_note_id  where note_id='"&note_id&"' order by note.note_time desc"
	rs.Open sql,conn,1,3
%>
<HTML>
<HEAD>
<TITLE>留言浏览</TITLE>
<meta http-equiv="Content-Type" content="text/html;	charset=gb2312">
<link rel="stylesheet" type="text/css" href="css/css.css">
</HEAD>
<BODY>
<%
	While(not rs.bof and not rs.Eof)
	note_title=rs("note_title")
	note_user=rs("note_user")
	note_content=rs("note_content")
%>
<table width="778px" align="center"
			cellpadding="0" cellspacing="0" bgcolor="" style="border:0;margin:10;align:center;">
  <tr height='32px' class='toptitle1' bgcolor='#ccddee'>
    <td height="27" colspan="2" bgcolor="#D9EBB9" style="padding-left:10px"><!--回复相对留言的缩进距离-->
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="93%" bgcolor="D9EBB9"><span class="style1">
		  <% if(rs("note_flag")=1) then
		  		response.Write("给版主的悄悄话...")
			else 
				response.write(note_title)
			end if
		   %></span></td>
          <td align="right"><img src="images/face/pic/<% =rs("note_user_pic")%>" width="24" height="24"></td>
        </tr>
    </table></td>
  </tr>
  <tr height='26px'>
    <!-- 标题行 -->
    <td colspan="2" bgcolor="#F9F8EF" style="padding-left:10px"><span class="style1">
	<% 
	if(rs("note_flag")<>1) then 
		response.write(note_user)
	%>：说
	<%end if%></span></td>
  </tr>
  <tr height='26px'>
    <!-- 内容行 -->
    <td colspan="2" bgcolor="#FFF6F8"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td valign="top" bgcolor="#F9F8EF" style="padding-left:25px; padding-right:10px;padding-top:6px;padding-bottom:2px;line-height:18px;width:710px">&nbsp;&nbsp;&nbsp;&nbsp;<%
		   if(rs("note_flag")=1) then
		   		response.Write("<img src='images/whisper.gif'>&nbsp;(给版主的悄悄话...)")
			else
				response.write(note_content)
			end if
			%> <br>
            <% if(note_flag=0 and note_answer=1) then %><br>
            <TABLE width="700" align="center" cellPadding=2 cellSpacing=1 bgcolor="#064727" class=embedbox>
              <TBODY>
                <TR>
                  <TD width="700" bgcolor="#FFFFFF" style="padding-top:6px; padding-left:10px; padding-bottom:2px; padding-right:10px;line-height:18px"><SPAN 
                  style="FONT-WEIGHT: bold; COLOR: #000000">&nbsp;版主回复：</SPAN> <SPAN style="COLOR: #000000">(<% =rs("noan_time")%>)</SPAN>&nbsp;
                    <hr color=#064727 size=1 style="width:700px; " >
                    <IMG 
                  src="images/face/pic/01.gif" width="90" height="90" class=face style="FLOAT: left; MARGIN: 2px 5px 5px 2px">&nbsp;&nbsp;<SPAN 
                  style="COLOR: #000000"><% =rs("noan_content")%></SPAN></TD>
                </TR>
              </TBODY>
            </TABLE>
			<% end if%>
          <br></td>
        </tr>
    </table></td>
  </tr>
  <tr class="bg_author">
    <!-- 留言人信息 -->
    <td align="right" style="padding-left:10px">留言时间：&nbsp;<% =rs("note_time") %>&nbsp;&nbsp;&nbsp;&nbsp;</td>
  </tr>
</table>
<%
	rs.MoveNext
	Wend
%>
</BODY>
</HTML>
