<!--#include file="conn/conn.asp"-->
<%
response.expires=0                '页面存活时间
response.CacheControl="no-cache"  '指定代理服务器是否缓冲由ASP产生的输出
%>
<% 
if session("flag")="登录" then 
'记录用户登录次数
sql_time="update Tab_user set accessTime=accessTime+1 where UserName='"&session("UserName")&"'"
conn.execute(sql_time)
'插入系统日志
sql_log = "insert into Tab_Log (UserName, Event) VALUES ('"&session("UserName")&_
"','[" &session("UserName")& "]登录本系统') "
conn.execute(sql_log)
session("flag")=""
end if
'判断Session变量是否过期
If Session("UserName")="" Then %>
<script language="javascript">
alert("您登录的网页已经过期，请重新登录！");
window.location="index.asp"
</script>
<%
End If
'查询用户信息
Set rs_userinfo = Server.CreateObject("ADODB.Recordset")
sql_UN = "select ID, UserName, purview, branch  FROM Tab_User  WHERE UserName = '"&_
session("UserName")&"'"
rs_userinfo.open sql_UN,conn,1,3
'查询公告栏信息
Set rs_bbs = Server.CreateObject("ADODB.Recordset")
sql_bbs = "SELECT ID, Subject FROM Tab_Placard ORDER BY dDate DESC"
rs_bbs.open sql_bbs,conn,1,3
%>
<style type="text/css">
<!--
body {
	margin-top: 0px;
	margin-bottom: 0px;
	margin-left: 0px;
	margin-right: 0px;
}
.STYLE5 {font-size: 9pt; color: #FFFFFF; }
a:link {
	color: #436283;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #436283;
}
a:hover {
	text-decoration: none;
	color: #436283;
}
a:active {
	text-decoration: none;
	color: #436283;
}
.STYLE8 {
	font-size: 9pt;
	color: #436283;
}
.STYLE9 {font-size: 10pt; color: #436283; }
-->
</style>

<script src="JS/DateTime2.js"></script>
<script language="javascript">
function Check()
{
if(confirm("确定要重新登录本系统吗??"))
{
window.location="index.asp";
}
}
</script>
<script language="javascript">
function Check1()
{
	window.close();
}
</script>
<table width="1005" height="80" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="1003" height="81" background="Images/top.gif"><table width="667" height="77" border="0" align="right" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="2"><table width="429" height="34"  border="0" align="right" cellpadding="-2" cellspacing="-2">
          <tr>
            <td width="5%" height="33" align="center" valign="middle"><img src="images/icon_smile.GIF" width="15" height="15" />&nbsp;</td>
            <td width="23%" align="center"><div align="center" class="STYLE5">用户名： <%=Session("UserName")%></div></td>
            <td width="28%" height="33"><div align="center" class="STYLE5"> 权限：<%=(rs_userinfo.Fields.Item("purview").Value)%></div></td>
            <td width="27%" height="33"><div align="center" class="STYLE5"> 所在部门：<%=(rs_userinfo.Fields.Item("branch").Value)%></div></td>
            <td width="17%" height="33"></td>
          </tr>
          <tr> </tr>
        </table></td>
      </tr>
      <tr>
        <td width="346" height="30" valign="bottom"><table width="335" border="0" cellpadding="0" cellspacing="0">
          <tr>
             <td width="229" height="24" align="right" class="STYLE9" id="Time"><script>ShowDate(Time);</script>
      &nbsp;&nbsp;&nbsp;</td>
    <td width="117" align="left" class="STYLE9" id="Time2">&nbsp;&nbsp;&nbsp;
        <script>ShowTime(Time2);</script></td>
              </tr>
        </table></td>
        <td width="321" valign="bottom"><table width="316" height="30" border="0" align="right" cellpadding="0" cellspacing="0">
          <tr>
            <td width="37" valign="bottom"><img src="Images/3.gif" width="32" height="30" /></td>
            <td width="48" class="STYLE8"><a href="default.asp" >首页</a></td>
            <td width="44" valign="bottom"><img src="Images/4.gif" width="32" height="30" /></td>
            <td width="64"><a href="#" class="STYLE8" onClick="Check();">重新登录</a></td>
            <td width="40" valign="bottom"><img src="Images/5.gif" width="32" height="29" /></td>
            <td width="83"><a href="#" class="STYLE8" onclick="Check1();">退出系统</a></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>

