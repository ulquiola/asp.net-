<!--#include file="top.asp"--><!--包含网站的头文件-->
<!-- #include file="Conn/conn.asp" --><!--数据库连接文件-->
<%
if request.QueryString("stype")="" then		'获取投票的类型
	if Request.ServerVariables("REMOTE_ADDR")=request.cookies("IPAddress") then
		'获取用户的IP地址，并判断该IP地址是否已经投过票
		response.write"<SCRIPT language=JavaScript>alert('感谢您的支持，您已经投过票了，请勿重复投票，谢谢！');"
		'弹出提示信息对话框
		response.write"history.back();</SCRIPT>"
		'返回到前一页面
	else
		options=request.form("Options")		'获取投票选项
		response.cookies("IPAddress")=Request.ServerVariables("REMOTE_ADDR") 	'获取IP地址
		conn.execute("update tb_Vote set Option"&options&"=Option"&options&"+1") '更新指定的记录
		conn.close																 '关闭记录集
		set conn=nothing														 '释放占用的系统资源空间
	end if
end if
select1="1.非常值得信赖"															'为变量赋值 
select2="2.比较可信"																'为变量赋值
select3="3.没什么感觉"																	
select4="4.完全不可信"
%>
<style type="text/css">
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
.STYLE3 {color: #000073}
</style>
<body>
<table width="800" height="526" border="0" align="center" cellpadding="0" cellspacing="0" background="images/new18.gif">
  <tr>
    <td height="86" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td width="46" height="289">&nbsp;</td>
    <td width="713" valign="top"><table width="677" height="70" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="677" height="70">
		
		<table width="677" height="65" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="7" height="5" background="images/00.gif">&nbsp;</td>
    <td width="697"><table width="648" border="0" cellspacing="0" cellpadding="0" align="center"  class="wenbenkuang">
<tr> 
<td width="648" height="29" class="table-xiayou"> 
<table border="0" cellpadding="0" cellspacing="0" width="90%" height="48" style="border-collapse: collapse" align="center">
<!-- #include file="Conn/conn.asp" -->     
<%
	set rs=server.CreateObject("adodb.recordset")		'创建记录集
	sql="select * from tb_Vote"							'查询数据
	rs.open sql,conn,1,3								'打开记录集
	while not rs.eof 									'判断是否有记录
%>
 <tr> 
<td height="48" valign="middle" colspan="4" align=center>&nbsp; <span class="STYLE3" style="font-size: 9pt">网站信任度投票结果</span></td>
   			  </tr>
        			<tr>
          				<td width="30%" valign="top" class="STYLE2">序号</td>
          				<td colspan="2" valign="top" class="STYLE2">百比分</td>
          				<td width="21%" valign="top" class="STYLE2">人数</td>
        			</tr>
<%
		an1=int(rs("Option1"))		'返回数字的整数部分
		an2=int(rs("Option2"))
		an3=int(rs("Option3"))
		an4=int(rs("Option4"))
		answer=an1+an2+an3+an4		'求出投票的总人数
		rs.movenext					'向下移动记录指针
	wend
%>
        			<tr valign="middle"> 
          				<td height="25" class="STYLE2"><%=select1%>：</td>
          				<td width="34%" height="25"><img src="images/RSCount.gif" width=<%=an1*100/answer%>% height=20><%'将图片按一定的百分比进行增加%></td>
       				  <td width="15%" height="25"><div align="center" class="STYLE2"><%=round(an1*100/answer,2)%>%</div><%'应用round函数返回按指定位数进行四舍五入的数值%></td>
          				<td class="STYLE2"><%=an1%>人</td>
	    			</tr>
		          	<tr valign="middle"> 
          				<td height="25" class="STYLE2"><%=select2%>：</td>
          				<td height="25"><img src=images/RSCount.gif width=<%=int(an2*100/answer)%>% height=20></td>
       				  <td height="25" class="STYLE2"><div align="center"><%=round(an2*100/answer,2)%>%</div><%'应用round函数返回按指定位数进行四舍五入的数值%></td>
          				<td class="STYLE2"><%=an2%>人</td>
					</tr>
		          	<tr valign="middle"> 
						<td height="25" class="STYLE2"><%=select3%>：</td>
          				<td height="25"><img src=images/RSCount.gif width=<%=int(an3*100/answer)%>% height=20></td>
       				  <td height="25"><div align="center" class="STYLE2"><%=round(an3*100/answer,2)%>%</div><%'应用round函数返回按指定位数进行四舍五入的数值%></td>
          				<td class="STYLE2"><%=an3%>人</td>
		  			</tr>
		          	<tr valign="middle"> 
          				<td height="25" class="STYLE2"><%=select4%>：</td>
          				<td height="25"><img src=images/RSCount.gif width=<%=int(an4*100/answer)%>% height=20></td>
       				  <td height="25" class="STYLE2"> <div align="center"><%=round(an4*100/answer,2)%>%</div></td>
          				<td class="STYLE2"><%=an4%>人</td>
		  			</tr>
        			<tr>
          				<td colspan="4" class="STYLE2"> 共有【<%=answer%>】人参加投票</td>
        			</tr>
		    </table>
      				<p align="center"><span class="STYLE2">【<a href="index.asp">返回首页</a>】</span> 
<%      
	set rs=nothing     
	conn.close     
	set conn=nothing
%>
</td>
</tr>
</table>
</td>
<td width="30" background="images/000.gif">&nbsp;</td>
</tr>
</table>
		</td>
      </tr>
    </table></td>
    <td width="41">&nbsp;</td>
  </tr>
  <tr>
    <td height="13" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td height="138" colspan="3">&nbsp;</td>
  </tr>
</table>