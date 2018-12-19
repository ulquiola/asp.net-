<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Request对象集合之：Form集合和QueryString集合</title>
<style>
body{
	margin:12px;
	font-size:12px;

}
</style>
</head>
<body>
<%'添加一个Form表单%>
<table width="399" height="166" border="0" cellpadding="0" cellspacing="0">
<form name="form1" method="post" action="index.asp">
  <tr>
    <td colspan="2"><div align="center">会员信息</div>
        <div align="center"></div></td>
  </tr>
  <tr>
    <td width="79"><div align="center">会员昵称：</div></td>
    <td width="203"><input name="name1" type="text" id="name1" value=<%=request.Form("name1")%> /></td>
  </tr>
  <tr>
    <td><div align="center">会员名称：</div></td>
    <td><input name="name2" type="text" id="name2" value="<%=request.Form("name2")%>" /></td>
  </tr>
  <tr>
    <td><div align="center">所属地区：</div></td>
    <td><select name="dq" id="dq">
      <option value="长春市">长春市</option>
      <option value="白城市">白城市</option>
      <option value="辽源市">辽源市</option>
      <option value="白山市">白山市</option>
      <option value="沈阳市">沈阳市</option>
    </select>
    </td>
  </tr>
  <tr>
    <td><div align="center">通讯地址：</div></td>
    <td><input name="address" type="text" id="address" value="<%=request.Form("address")%>" /></td>
  </tr>
  <tr>
    <td colspan="2"><div align="center">
      <input type="submit" name="Submit" value="提交" />
      <%'添加一个提交按钮%>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <input type="reset" name="Submit2" value="重置" />
      <%'添加一个重置按钮%>
    </div></td>
  </tr>
</form>
</table>

<hr />
<%
	if request.Form("name1") <> "" then
%>
&nbsp;&nbsp;&nbsp;&nbsp;会员信息<br />
会员昵称：<%=request.Form("name1")%><br> 	<%'应用Form数据集合获取表单数据%>
会员名称：<%=request.Form("name2")%><br>	<%'应用Form数据集合获取表单数据%>
所属地区：<%=request.Form("dq")%><br> 		<%'应用Form数据集合获取表单数据%>
通讯地址：<%=request.Form("address")%> 		<%'应用Form数据集合获取表单数据%>
<%
	end if
%>


</body>
</html>
