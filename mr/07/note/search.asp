<!--#include file="common/conn.asp"-->
<%
	key_words = request.Form("key_words")
	Set rs = Server.CreateObject("ADODB.RecordSet")
	'����������Ϣ
	sql = "select note.*,noan.* from tb_note as note left join tb_note_answer as noan on note.note_id = noan.noan_note_id where note.note_title like '%"&key_words&"%' or note.note_content like '%"&key_words&"%' or note.note_time like '%"&key_words&"%' or note.note_user like '%"&key_words&"%'"
	rs.Open sql,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>���Լ���</TITLE>
<META NAME="Author">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="css/css.css">
</HEAD>
<BODY>
<table width="778" align="center" cellpadding="0" cellspacing="0" bgcolor="#eeeeee" style="border:0;margin:10;align:center;">
	<tr height='32px' class='toptitle1' bgcolor='#ccddee'>
	<td width='10'></td>
	<td height="24" colspan=2 bgcolor="#D9EBB9" style="padding-left:10px">
	<span class="style1">վ������:<% response.Write(key_words) %></span>	</td>
  </tr>
  <tr height='32px'>
	<td height="25"></td>
	<td colspan=2 bgcolor="#FFFFFF">
	&nbsp;&nbsp;��������б�	</td>
  </tr>
	<tr><td></td>
	<td height="24" valign="top" bgcolor="#FFFFFF">
	<%
	While(Not rs.Eof and not rs.bof)
	note_user=rs("note_user")
	note_title=rs("note_title")
	note_content=rs("note_content")
	//����������������Ϣ
	%>
	<table width="778px" align="center"	cellpadding="0" cellspacing="0" bgcolor="" style="border:0;align:center;">
      <tr height='32px' class='toptitle1' bgcolor='#ccddee'>
        <td height="27" colspan="2" bgcolor="#D9EBB9" style="padding-left:10px"><!--�ظ�������Ե���������-->
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="100%" bgcolor="#D9EBB9"><span class="style1">
                  <% 
				  if(rs("note_flag")=1) then 
				  	 response.write("�����������Ļ�...") 
				  else
				     response.write(note_title)
				  end if 
				  %>
                </span></td>
                <td align="right"><img src="images/face/pic/<% =rs("note_user_pic") %>" width="24" height="24"></td>
              </tr>
          </table></td>
      </tr>
      <tr height='26px'>
        <!-- ������ -->
        <td colspan="2" bgcolor="#F9F8EF" style="padding-left:10px"><span class="style1">
          <% if(rs("note_flag")<>1) then 
		       	response.write(note_user)
		  %> ��˵
          <% end if %>
        </span></td>
      </tr>
      <tr height='26px'>
        <!-- ������ -->
        <td colspan="2" bgcolor="#F9F8EF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td valign="top" style="padding-left:25px; padding-right:10px;padding-top:6px;padding-bottom:2px;line-height:18px;width:710px">
                  <%
				   if(rs("note_flag")=1) then
				   		response.Write("<img src='images/whisper.gif'>&nbsp;(�����������Ļ�...)")
				   else
					    response.write(note_content)
				   end if	
				   %>
                  <br>
                  <% if rs("note_flag")=0 and rs("note_answer")=1 then %>
                  <br>
                  <TABLE width="700" align="center" cellPadding=2 cellSpacing=1 bgcolor="#D9D2B6" class=embedbox >
                    <TBODY>
                      <TR>
                        <TD width="700" bgcolor="#FFFFFF" style="padding-top:6px; padding-left:10px; padding-bottom:2px; padding-right:10px;line-height:18px"><SPAN 
                  style="FONT-WEIGHT: bold; COLOR: #000000">&nbsp;�����ظ���</SPAN> <SPAN style="COLOR: #000000">(<% =rs("noan_time") %>)</SPAN>&nbsp;
                            <hr color=#D9D2B6 size=1 style="width:700px; " >
                        <IMG 
                  src="images/face/pic/01.gif" width="90" height="90" class=face style="FLOAT: left; MARGIN: 2px 5px 5px 2px">&nbsp;&nbsp;<SPAN 
                  style="COLOR: #000000"><% =rs("noan_content") %></SPAN> </TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                <% end if %>
                  <br></td>
            </tr>
        </table></td>
      </tr>
      <tr class="bg_author">
        <!-- ��������Ϣ -->
        <td align="right" style="padding-left:10px">����ʱ�䣺&nbsp;<% =rs("note_time") %>&nbsp;&nbsp;&nbsp;&nbsp;</td>
      </tr>
    </table></td>
<%
	rs.MoveNext
	Wend
%>
</table>
</BODY>
</HTML>
