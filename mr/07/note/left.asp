<link href="css/css.css" rel="stylesheet" type="text/css" />
	<table width="199" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="741" valign="top">
          <table width="199" height="80" border="0" align="center" cellpadding="0" cellspacing="1">
            <tr>
              <td height="80" align="center" valign="top" background="images/so_line.jpg" bgcolor="#FFFFFF"><table width="199" height="52" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td background="images/so.jpg">&nbsp;&nbsp;&nbsp;</td>
                  </tr>
                </table>
				<br />
                  可按标题、内容、姓名、时间检索
				  <form target="_blank" method="post" action="search.asp">
                    <table width="180" height="40" border="0" align="center" cellpadding="0" cellspacing="1">
                      <tr>
                        <td align="center" valign="middle" bgcolor="#FFFFFF"><input name="key_words" type="text" class="btn1" value=" 搜索关键字" size='19' />&nbsp;
                        <input name="submit" type="submit" class="btn1" value="检索" />
						</td>
					</tr>
                    </table>
              </form></td>
            </tr>
          </table>
          <table width="199" height="24" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="199" background="images/so_bottom.jpg"></td>
            </tr>
          </table>
          <table width="199" height="93" border="0" align="center" cellpadding="0" cellspacing="1">
            <tr>
              <td height="91" valign="top" bgcolor="#FFFFFF"><table width="199" height="24" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td height="37" background="images/leave_line.jpg">&nbsp;</td>
                  </tr>
                </table>
                  <table width="199" height="32" border="0" align="center" cellpadding="0" cellspacing="0">
				  <!--#include file="common/conn.asp"-->
                    <%
					   Set rs1 = Server.CreateObject("ADODB.RecordSet")
					   sql1="select top 5  note_title ,note_id ,note_time from tb_note order by note_id desc"
					   rs1.Open sql1,conn,1,3
					   if rs1.bof and rs1.eof then
					%>
                    <tr>
                      <td height="22"><div align="center">暂无留言！</div></td>
                    </tr>
                    <%
					  else
						While(not rs1.Eof)
					%>
                    <tr>
                      <td height="32" background="images/leave_line2.jpg"> &nbsp;&nbsp;&nbsp;&nbsp;<a href="user_note_read.asp?id=<% =rs1("note_id")%>">
             	 <% 
					If(Len(rs1("note_title")) >20) Then
						Response.Write(Server.HTMLEncode(Mid(rs1("note_title"),1,20))&"....")
					Else
						Response.Write(Server.HTMLEncode(rs1("note_title")))
					End If
				  %>
                      </a></td>
                    </tr>
				  <%
					rs1.MoveNext
					Wend
				  end if
				  %>
              </table></td>
            </tr>
          </table>
          <table width="199" height="35" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td background="images/sm.jpg"></td>
            </tr>
          </table>
          <table width="199" height="230" border="0" align="center" cellpadding="0" cellspacing="1">
            <tr>
              <td height="228" valign="top" bgcolor="#FFFFFF"><table width="100%" height="44" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td bgcolor="F1F0ED" style=" padding-left:4px;padding-top:12px;padding-right:0px;padding-bottom:12px;">请自觉遵守本站声明的以下条款：</td>
                    </tr>
                    <tr>
                      <td height="30" bgcolor="F1F0ED" style="padding-left:6px;line-height:20px"><img src="images/a_1.gif" width="16" height="9" />&nbsp;尊重网上道德，遵守中华人民共和国的各项有关法律法规</td>
                    </tr>
					<tr>
                      <td height="30" bgcolor="F1F0ED" style="padding-left:6px;line-height:20px"><img src="images/a_1.gif" width="16" height="9" />&nbsp;本站管理员有权保留或删除其留言中的任意内容</td>
                    </tr>
					<tr>
                      <td height="30" bgcolor="F1F0ED" style="padding-left:6px;line-height:20px"><img src="images/a_1.gif" width="16" height="9" />&nbsp;本站禁止在留言中发表敏感词</td>
                    </tr>
					<tr>
                      <td height="30" bgcolor="F1F0ED" style="padding-left:6px;line-height:20px"><img src="images/a_1.gif" width="16" height="9" />&nbsp;本站有权在网站内转载或引用您的留言</td>
                    </tr>			
					<tr>
                      <td height="30" bgcolor="F1F0ED" style="padding-left:6px;line-height:20px"><img src="images/a_1.gif" width="16" height="9" />&nbsp;参与本留言表明您已经阅读并接受上述条款</td>
                    </tr>	
              </table></td>
            </tr>
          </table>
          <table width="170" height="34" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="34"><img src="images/a.gif" width="170" height="42" /></td>
            </tr>
          </table>
          <br />
          <table width="170" height="34" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="34"><img src="images/c.gif" width="170" height="42" /></td>
            </tr>
          </table>
          <br />
          <table width="170" height="34" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="34"><img src="images/b.gif" width="170" height="42" /></td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">