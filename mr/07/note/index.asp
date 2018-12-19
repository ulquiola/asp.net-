<!--#include file="common/conn.asp"-->
<!--#include file="Replace.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/css.css" rel="stylesheet" type="text/css">
<title>美丽心情留言本</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	text-align:center

}
.STYLE22 {color: #CC3333}
-->
</style>
</head>
<body>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="3"><!--#include file="header.html"--></td>
  </tr>
  <tr>
    <td width="190" height="522"  valign="top" bgcolor="#FFFFFF"><!--#include file="left.asp"--></td>
    <td width="3" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
    <td width="757" valign="top" bgcolor="#F9F8EF"><table width="755" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center" valign="top" scope="col"><table width="755" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="755" height="47" background="images/lyb.jpg" scope="col">&nbsp;</td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td><br>
                  <table width="100%"  border="0" cellpadding="2" cellspacing="2" bgcolor="#D9EBB9">
                    <tr>
                      <th align="left" bgcolor="#FFFFFF" scope="col">&nbsp;<img src="images/notice.gif" width="20" height="20"><span class="style12">贴心提醒：</span> </th>
                    </tr>
                    <tr>
                      <td height="40" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;欢迎光临美丽心情留言本，希望您在这里渡过开心快乐的每一刻，留下您珍贵的一笔是对本站最大的帮助，谢谢！版主---纯净水</td>
                    </tr>
                  </table></td>
              </tr>
              <tr valign="top">
                <td align="center">
				<%
				sql = "select tb_note.*,answ.* from tb_note left join (select noan_note_id,noan_content,noan_time from tb_note_answer) as answ on answ.noan_note_id = tb_note.note_id  order by note_time desc "
				Set rs = Server.CreateObject("ADODB.RecordSet")
				rs.Open sql,conn,1,3
				rs.PageSize = 3
				'实现分页
				if rs.Eof then
					rs_total = 0
				else
					rs_total = rs.RecordCount
				end if
				dim pageno
				getpageno = Switch(Request("pageno"))
				if(getpageno = "")then
					pageno = 1
				else
					pageno = getpageno
				End if
				if(not rs.Eof)then
					rs.AbsolutePage = pageno
				end if
				%>	
				<%
				repeat_rows = 0
				While(repeat_rows < rs.PageSize and Not rs.Eof)
				title=rs("note_title")
				note_content=rs("note_content")
				%>
                  <br>
                  <table width="100%"  border="0" cellpadding="2" cellspacing="1" bgcolor="#205401">
                    <tr bgcolor="#F9F8EF">
                      <td rowspan="2" align="center" valign="top" scope="col" width="19%"><br>
                       <%=rs("note_user")%><br>
                        <img src="images/face/pic/<%=rs("note_user_pic")%>" width="90" height="90"><br><br>
                        <img src="images/time.jpg" width="15" height="15">&nbsp;<%=rs("note_time")%><br><br>
					  </td>
                      <th width="81%" height="24" align="left" scope="col">&nbsp;&nbsp;
                        <% 
						if rs("note_flag")=1 then
							response.Write("给版主的悄悄话...")
						else
							response.Write(title)
						end if
						%>
					  </th>
                    </tr>
                    <tr bgcolor="#F9F8EF">
                      <td height="113" align="left" valign="top" bgcolor="#F9F8EF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                            <td height="39" style="padding-left:10px; padding-right:10px; line-height:18px">&nbsp;&nbsp;&nbsp;&nbsp;
                              <% 
							  if rs("note_flag")=1 then
							  		response.Write( "<img src='images/whisper.gif'>&nbsp; 给版主的悄悄话...")
							  else
							  		response.Write(note_content)
							  end if
							   %>
                              <br></td>
                          </tr>
                          <tr>
                            <td align="center"><TABLE 
      style="BORDER-TOP-WIDTH: 0px; BORDER-LEFT-WIDTH: 0px; BORDER-BOTTOM-WIDTH: 0px; WIDTH: 100%; BORDER-COLLAPSE: collapse; BORDER-RIGHT-WIDTH: 0px" 
      cellSpacing=0 cellPadding=2>
                              </TABLE>
                           <%
						   if rs("note_flag")=0 and rs("note_answer")=1 then
						   %>
                                <TABLE width="500" cellPadding=2 cellSpacing=1 bgcolor="#205401" >
                                  <TBODY>
                                    <TR>
                                      <TD width="500" bgcolor="#FFFFFF" style="padding-top:10px; padding-left:10px; padding-bottom:10px; padding-right:10px;line-height:18px"><SPAN style="FONT-WEIGHT: bold; COLOR: #000000">&nbsp;版主回复：</SPAN> <SPAN style="COLOR: #000000">(<% =rs("noan_time")%>)</SPAN>&nbsp;
                                       <hr color=#205401 size=1 style="width:500px;">
                                      <IMG 
                  src="images/face/pic/01.gif" width="90" height="90" class=face style="FLOAT: left; MARGIN: 2px 5px 5px 2px">&nbsp;&nbsp;<SPAN 
                  style="COLOR: #000000"><% =rs("noan_content")%></SPAN> </TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <br>
                           <% end if %></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table>
				<%
				repeat_rows = repeat_rows + 1
				rs.MoveNext
				Wend
				%>
              <tr>
                <td><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#205401">
                    <tr height='26px' align="right">
                      <th align="center" bgcolor="#F9F8EF">
                        &nbsp;&nbsp; 『留言分页』</th>
                    </tr>
                    <tr height='26px'>
                      <td align="center" bgcolor="#F9F8EF">
					  <table cellspacing="0" cellpadding="0" border="0" height="24" width="755" align="center">
                        <tr>
                          <td align="right" nowrap> 每页『<span class="STYLE22">&nbsp;<%=rs.PageSize%>&nbsp;</span>』条&nbsp;/&nbsp;共『<span class="STYLE22">&nbsp;<%=rs.RecordCount%>&nbsp;</span>』条&nbsp;当前『&nbsp;<span class="STYLE22"><%=pageno%>&nbsp;</span>』页/共『<span class="STYLE22">&nbsp;<%=rs.PageCount%>&nbsp;</span>』页&nbsp;&nbsp;&nbsp; 
					<%
					if(pageno <> 1) then
					%>
                            &nbsp;<a href="index.asp?pageno=1">第一页</a>
                    <%
					End if
					if(pageno <> 1)then
					%>
                            <a href="?pageno=<%=(pageno-1)%>">上一页</a>
                    <%
					end if
					if(instr(pageno,cstr(rs.pagecount)) = 0) then
					%>
                            <a href="?pageno=<%=(pageno+1)%>">下一页</a>
                    <%
					end if
					if(instr(pageno,cstr(rs.pagecount)) = 0) then
					%>
                            <a href="?pageno=<%=rs.pagecount%>">最后一页</a>
                    <%
					end if
					rs.close
					Set rs = Nothing
					%>
&nbsp;&nbsp; </td>
                        </tr>
                      </table></td>
                    </tr>
                  </table>
			    </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td colspan="3"><!--#include file="footer.html"--></td>
  </tr>
</table>
</body>
