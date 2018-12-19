<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"-->
<!--#include file="inc/customFunc.asp"-->
<%
set rs=server.CreateObject("ADODB.RecordSet")
sql="select * from tb_board order by id desc"
rs.open sql,conn,1,3
person=replace(request.Form("person"),"'","''")
if person<>"" then
	tel=replace(request.Form("tel"),"'","''")
	qq=replace(request.Form("qq"),"'","''")
	email=replace(request.Form("email"),"'","''")
	yuyueren=replace(request.Form("yuyueren"),"'","''")
	talktime=replace(request.Form("talktime"),"'","''")
	title=replace(request.Form("title"),"'","''")
	content=replace(request.Form("content"),"'","''")
	on error resume next
	sql="insert into tb_board (person,tel,qq,email,yuyueren,talktime,title,content) values('"&person&"','"&tel&"','"&qq&"','"&email&"','"&yuyueren&"','"&talktime&"','"&title&"','"&content&"')"
	conn.execute(sql)
	if err<>0 then
		call message("留言失败!","board.asp")
	else
		call message("留言成功!","board.asp")
	end if
end if
%>
<html>
<head>
<title>留言板</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="CSS/style.css" rel="stylesheet">
<script language="javascript">
function check(Form){
	for(i=0;i<Form.length;i++){  
		if(Form.elements[i].value == ""){         //Form的属性elements的首字e要小写
			alert(Form.elements[i].title + "不能为空!");
			Form.elements[i].focus();
			return false;
		}
	}
}
</script>
</head>
<body>
<table width="625"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bordercolordark="#1985DB" bordercolorlight="#FFFFFF">
	<tr>
		<td colspan="2" height="25px" align="center">&nbsp;留言列表</td>
	</tr>
  <tr>
    <td width="173" valign="top" bgcolor="#FFFFFF">
		<%if rs.eof and rs.bof then%>
		<div style=" text-align:center;">暂无留言信息!</div>
		<% else %>
		<table border="0" cellpadding="0" cellspacing="0">
		<%
			for i=1 to rs.recordcount
		%>	
			<tr>
				<td><a href="board_detail.asp?ID=<%=rs("id")%>"><img src="images/boardlist.ico" width="20" height="20" border="0"> <%=server.HTMLEncode(rs("title"))%></a></td>
			</tr>
		<%
			rs.movenext
			next
		%>
		</table>
		<%
			end if
		%>
	</td>	
    <td width="446" valign="top" bgcolor="#FFFFFF">
		<table width="109%" height="238"  border="0" cellpadding="0" cellspacing="0" class="tableBorder">
			<form name="form1" method="post" action="#">
                          
                            <tr>
                              <td width="16%" height="25" align="center" valign="middle">留 言 人： </td>
                              <td colspan="3" valign="middle"><input name="person" type="text" class="txt_grey" size="18" title="留言人">
                                * </td>
                            </tr>
                            <tr>
                              <td width="16%" height="25" align="center" valign="middle">联系电话：</td>
                              <td width="28%" valign="middle"><input name="tel" type="text" size="17" title="联系电话">
                              *</td>
                              <td width="21%" align="center" valign="middle">QQ号码：</td>
                              <td width="35%" valign="middle"><input name="qq" type="text" size="16" title="QQ">
                              *</td>
                            </tr>
                            <tr>
                              <td height="25" align="center" valign="middle">邮箱地址：</td>
                              <td colspan="3" valign="middle"><input name="email" type="text" size="35" title="邮箱地址">
                              *</td>
                            </tr>
                            <tr>
                              <td height="25" align="center" valign="middle">预&nbsp;约&nbsp;人：</td>
                              <td width="28%" valign="middle"><input name="yuyueren" type="text" size="17" title="预约人">
                              *</td>
                              <td width="21%" align="center" valign="middle">预约时间段：</td>
                              <td width="35%" valign="middle"><input name="talktime" type="text" size="16" title="谈话时间段">
                              *</td>
                            </tr>
                            <tr>
                              <td height="25" align="center">咨询主题：</td>
                              <td colspan="4"><input name="title" type="text" class="txt_grey" size="72" title="留言主题">
                                *</td>
                            </tr>
                            <tr>
                              <td height="25" align="center">咨询事宜：</td>
                              <td colspan="4"><textarea name="content" cols="40" rows="8" class="textarea" title="留言内容"></textarea>
                                *</td>
                            </tr>
                            <tr>
                              <td height="25" colspan="5" align="center"><input name="Submit" type="submit" class="btn_grey" value="保存" onClick="return check(form1)">
                                &nbsp;
                                <input name="Submit2" type="reset" class="btn_grey" value="重置"></td>
                            </tr>
							</form>
                          </table>	
	</td>
  </tr>
</table>
</body>
</html>
<%
closeRS(rs)
closeconn(conn)
%>
