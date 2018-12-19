<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../conn/conn.asp"-->
<!--#include file="../inc/customFunc.asp"-->
<%
if request.QueryString("ID")<>"" then
ID=request.QueryString("ID")
end if
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from tb_board where ID="&ID
rs.open sql,conn,1,3
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<script language="javascript">
function check(myform)
{
	if (myform.content1.value=="")
	{
		alert("请输入回复内容!");
		myform.content1.focus();
		return false;
	}
}
</script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
-->
</style></head>
<body>
<table width="403" height="230" align="center" cellpadding="0" cellspacing="0">
<form name="myform" method="post" action="reply.asp" />
  <tr>
    <td width="401" height="208" valign="top"><table width="430" height="181" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="12" colspan="2" bgcolor="#327387"><div align="center" class="STYLE2"><span class="STYLE1"><font color="#FFFFFF">留言簿信息查看页面</font></span></div></td>
      </tr>
      <tr>
        <td width="87" height="8"><div align="center"><span class="STYLE2">留言人：</span></div></td>
        <td width="343"><input name="person" type="text" id="person" value="<%=rs("person")%>" readonly="readonly" /></td>
      </tr>
      <tr>
        <td width="87" height="8"><div align="center"><span class="STYLE2">联系电话：</span></div></td>
        <td width="343"><input name="person" type="text" id="person" value="<%=rs("tel")%>" readonly="readonly" /></td>
      </tr>
      <tr>
        <td width="87" height="8"><div align="center"><span class="STYLE2">QQ：</span></div></td>
        <td width="343"><input name="person" type="text" id="person" value="<%=rs("qq")%>" readonly="readonly" /></td>
      </tr>
      <tr>
        <td width="87" height="8"><div align="center"><span class="STYLE2">邮箱：</span></div></td>
        <td width="343"><input name="person" type="text" id="person" value="<%=rs("email")%>" readonly="readonly" /></td>
      </tr>
      <tr>
        <td width="87" height="8"><div align="center"><span class="STYLE2">预约人：</span></div></td>
        <td width="343"><input name="person" type="text" id="person" value="<%=rs("yuyueren")%>" readonly="readonly" /></td>
      </tr>
      <tr>
        <td width="87" height="8"><div align="center"><span class="STYLE2">谈话时间段：</span></div></td>
        <td width="343"><input name="person" type="text" id="person" value="<%=rs("talktime")%>" readonly="readonly" /></td>
      </tr>
      <tr>
        <td height="17"><div align="center" class="STYLE2">咨询主题：</div></td>
        <td><input name="title" type="text" id="title" value="<%=rs("title")%>" size="30" readonly="readonly" /></td>
      </tr>
      <tr>
        <td height="30"><div align="center" class="STYLE2">咨询事宜：</div></td>
        <td><textarea name="textarea" cols="30" rows="5" readonly="readonly" ><%if rs("reply")<>"" then%><%=rs("content")%><%else%>暂无回复信息！<%end if%>
        </textarea></td>
      </tr>
      <tr>
        <td height="30"><div align="center" class="STYLE2">回复留言：</div></td>
        <td><textarea name="content1" cols="30" rows="5" ><%=rs("reply")%></textarea></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="20"><div align="center">
	<Input name="ID" type="hidden" value="<%=rs("id")%>" /><input type="submit" value="回复" style="border: 0px; background-color:#FFFFFF;" onClick="return check(myform)" />&nbsp;&nbsp;<input  type="button" value="关闭窗口" onClick="javascript:window.close()"  style="border: 0px; background-color:#FFFFFF;"></div></td>
  </tr>
</table>
</body>
</html>