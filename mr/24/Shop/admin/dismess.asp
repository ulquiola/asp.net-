<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="chk_admin.asp" -->
<!--#include file="../include/include.asp" -->
<script language = "JavaScript">
function adddismess()
{
	if(document.myform.user.value=="") 
	{
		document.myform.user.focus();
		alert("请输入发表人！");
		return false;
	}
	if(document.myform.shijian.value=="") 
	{
		document.myform.shijian.focus();
		alert("请输入时间！");
		return false;
	}
	if(document.myform.biaoti.value=="") 
	{
		document.myform.biaoti.focus();
		alert("请输入标题！");
		return false;
	}
	if(document.myform.neirong.value=="") 
	{
		document.myform.neirong.focus();
		alert("请输入内容！");
		return false;
	}
}
</script>
<%
if request("action")="del" then
	conn.execute("delete from [fankui]  where id="&request("id")&"")
	response.Write("<script>alert('删除成功！');window.location.href='dismess.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">

  <tr> 
    <td align="center"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <form action="dismess.asp" name="form" method="post">
	  <tr>
        <td width="46%">&nbsp;</td>
        <td width="29%"><font color="#FFFFFF">查看意见反馈</font></td>
        <td width="25%">
<!-- 当对下拉选择进行操作时利用脚本提交表单 -->
<select name="leixing" onChange="this.form.submit();">
<!-- end -->
  <option value="1">对网站的建议</option>
  <option value="2">对公司的建议</option>
  <option value="3">对产品的投诉</option>
  <option value="4">对服务的投诉</option>
</select>	
		</td>
      </tr>
	 </form>
    </table></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="128"><font color="#FF0000">发表人</font></td>
          <td width="420"><font color="#FF0000">标题名</font></td>
          <td width="153"><font color="#FF0000">时 间</font></td>
          <td width="85"><font color="#FF0000">操  作</font></td>
        </tr>
<%
if  request("leixing")<>"" then
	leixing="where leixing='"&request("leixing")&"'"	'根据接收到的值来设置 SQL 语句的执行条件
else
	leixing="where leixing='1'"
end if
sql="select * from [fankui] "&leixing&" order by id desc;"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof And rs.bof then
	Response.Write "<p align='center' class='contents'> 对不起，暂无内容！</p>"
else
	rs.pagesize=10
	SafeRequest(request("page"))
	page=clng(request("page"))
	if page<1 then page=1
	if page>rs.pagecount then page=rs.pagecount
	show rs,page
	sub show(rs,page)
	rs.absolutepage=page
	for i=1 to rs.pagesize
%>
        <tr bgcolor="#FFFFFF" align="center"> 
            <td><%=rs("user")%></td>
            <td><a href="#" onClick='javascript:window.open("lookdismess.asp?id=<%=rs("id")%>")'><%=rs("biaoti")%></a></td>
            <td><%=rs("shijian")%></td>
            <td><a href="dismess.asp?action=del&id=<%=rs("id")%>">删除</a></td>
        </tr>
<%
	rs.movenext
	if rs.eof then exit for
	next
	end sub
%> 
      <tr>
	  <form action='' method='get' name='form'>
        <td colspan="6" bgcolor="#FFFFFF">
          <div align="center">
<%	
	if page<>1 then
		response.Write("&nbsp;&nbsp;<a href="&path&"?page=1>第一页</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&(page-1)&" >上一页</a>")
	end if 
	response.Write("&nbsp;&nbsp;当前 <font color='#FF0000'>"&page&"</font> 页/<font color='#FF0000'>"&rs.pagecount&"</font> 页")
	response.Write("&nbsp;&nbsp;共 <font color='#FF0000'>"&rs.recordcount&"</font> 条")
	if page<>rs.pagecount then
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&(page+1)&">下一页</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&rs.pagecount&">最末页</a>")
	end if
	response.Write("&nbsp;&nbsp;跳转到<input type='text' size='2' name='page'>页<input type='submit' value='GO'>")
end if
rs.close
set rs=nothing
%>
		    </div></td>
        </form></tr>
      </table>
      <br></td>
  </tr>
</table>