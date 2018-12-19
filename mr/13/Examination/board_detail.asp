<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"-->
<!--#include file="inc/customFunc.asp"-->
<%
set rs=server.CreateObject("ADODB.RecordSet")
sql="select * from tb_board"
rs.open sql,conn,1,3
ID=request.QueryString("ID")
if ID<>"" then
	set rs_B=server.CreateObject("ADODB.RecordSet")
	sql="select * from tb_board where ID="&ID
	rs_B.open sql,conn,1,3
else
	call message("您的操作有误!","board.asp")
end if
%>
<html>
<head>
<title>留言板</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="CSS/style.css" rel="stylesheet">
</head>
<body> 
<table width="634"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bordercolordark="#1985DB" bordercolorlight="#FFFFFF"> 
<tr> 
          <td height="26" colspan="2" align="center">&nbsp;留言列表</td> 
    </tr> 
  <tr> 
    <td width="138" height="296" valign="top" bgcolor="#FFFFFF">
		<table width="100%" height="25"  border="0" cellpadding="1" cellspacing="1" bgcolor="#FFFFFF">
            <%
			 for i=1 to rs.recordcount
			 %>
            <tr>
            	<td><a href="board_detail.asp?ID=<%=rs("id")%>"><img src="Images/boardlist.ico" width="20" height="20" border="0"><%=server.HTMLEncode(rs("title"))%></a></td>
            </tr>
            <%
			 	rs.movenext
			 next
			 %>
    </table>
	</td> 
    <td width="490" valign="top" bgcolor="#FFFFFF">
		<table width="100%" cols="5" border="0" cellpadding="0" cellspacing="0">
			<tr><td colspan="5" align="center" height="15px">&nbsp;</td></tr>
        	<tr>
				<td width="19%" height="29" align="center" valign="middle" class="word_blue">留 言 人： </td>
				<td width="29%" valign="middle"><%=server.HTMLEncode(rs_B("person"))%></td> 
				<td width="39%" valign="middle">&nbsp;</td>
				<td width="13%" valign="middle">&nbsp;</td>
			</tr>
			<tr>
				<td width="19%" height="25" align="center" class="word_blue">联系电话：</td>
				<td><%=server.HTMLEncode(rs_B("tel"))%></td>
				<td align="left">&nbsp;QQ号码：<%=server.HTMLEncode(rs_B("qq"))%></td>
				<td colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td height="25" align="center" class="word_blue">邮箱地址：</td>
				<td colspan="4"><%=server.HTMLEncode(rs_B("email"))%></td> 
            </tr> 
            <tr>
				<td height="25" align="center" class="word_blue">预&nbsp;约&nbsp;人：</td>
				<td width="29%"><%=server.HTMLEncode(rs_B("yuyueren"))%></td>
				<td width="39%" align="left">&nbsp;谈话时间段：<%=server.HTMLEncode(rs_B("talktime"))%></td>
				<td colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td height="25" align="center" class="word_blue">留言主题：</td> 
				<td colspan="4"><%=server.HTMLEncode(rs_B("title"))%></td> 
			</tr> 
			<tr> 
				<td height="28" align="center" class="word_blue">留言内容：</td> 
				<td colspan="4" class="word_orange">&nbsp;</td> 
			</tr>
			<tr valign="top">
				<td height="94" colspan="5" style="padding:5px;"> <%=convert(server.HTMLEncode(rs_B("content")))%></td>
			</tr>
			<%if not isnull(rs_B("reply")) then%>
			<tr>
				<td colspan="5"><hr class="word_gray" size="1" width="96%"></td>
			</tr>
            <tr>
				<td height="43" align="center" class="word_blue">回复内容：</td>
				<td height="43" colspan="4" class="word_orange"><%=rs_B("reply")%></td>
			</tr>
			<%
				end if
				closeRS(rs)
				closeconn(conn)
			%>
		</table> 
	</td> 
  </tr> 
</table> 
</body>
</html>
