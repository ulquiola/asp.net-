<!--#include file="Conn/conn.asp" -->
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
sql_log = "INSERT INTO Tab_Log (UserName, Event) VALUES ('"&session("UserName")&_
"','[" &session("UserName")& "]登录本系统') "
conn.execute(sql_log)
session("flag")=""
end if
'判断Session变量是否过期
If Session("UserName")="" Then %>
<script language="javascript">
alert("您登录的网页已经过期，请重新登录！");
window.location="login.asp"
</script>
<%
End If
'查询用户信息
Set rs_userinfo = Server.CreateObject("ADODB.Recordset")
sql_UN = "SELECT ID, UserName, purview, branch  FROM Tab_User  WHERE UserName = '"&_
session("UserName")&"'"
rs_userinfo.open sql_UN,conn,1,3

Set rs_email = Server.CreateObject("ADODB.Recordset")
sql_email="SELECT ID,subject,SPerson,flag,sTime FROM Tab_Send where lperson='"&session("username")&"'"
rs_email.open sql_email,conn,1,3
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<style type="text/css">
<!--
body {
	margin: 0px;
}
.STYLE1 {
	font-size: 10pt;
	color: #000000;
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
</style>
<table width="814" height="520" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" background="Images/main.gif"><table width="719" height="480" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="719" height="59" valign="bottom"><table width="196" height="30" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="38"><div align="right"><img src="Images/email.gif" width="32" height="33" /></div></td>
            <td width="158" valign="bottom">&nbsp;<img src="Images/ess.gif" width="68" height="20" /></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td valign="middle"><table width="661" height="312" border="0" cellpadding="-2" cellspacing="-2">
          <tr>
            <td width="661" valign="top" background="images/default/default_03.gif" ><% If rs_email.EOF And rs_email.BOF Then %>
                <table width="100%"  border="0" cellspacing="-2" cellpadding="-2">
                  <tr>
                    <td height="63"><div align="center" class="style4 style5 STYLE1"> 暂时没有任何文件！</div></td>
                  </tr>
                </table>
              <% else %>
                <table width="100%" height="27"  border="1" cellpadding="-2" cellspacing="-2"
	   bordercolor="#FFFFFF" bordercolorlight="#FFCCCC" bordercolordark="#FFFFFF">
                  <tr>
                    <td width="10%" height="21"><div align="center"><span class="STYLE1">状态 </span>&nbsp;</div></td>
                    <td width="43%"><div align="center" class="STYLE1">文件主题</div></td>
                    <td width="13%"><div align="center" class="STYLE1">发件人</div></td>
                    <td width="26%"><div align="center" class="STYLE1">时间</div></td>
                    <td width="8%"><div align="center" class="STYLE1">删除</div></td>
                  </tr>
                  <%'分页'
	rs_email.pagesize=12
	page1=CLng(Request("page1"))
	if page1<1 then page1=1
		rs_email.absolutepage=page1
		for i=1 to rs_email.pagesize %>
                  <tr>
                    <td height="21"><div align="center">
                        <!--标识邮件的读取状态-->
                        <% if rs_email("flag")="否" then %>
                        <!--播放Flash动画-->
                        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" 
			  codebase="http://download.macromedia.com/pub/shockwave/cabs/flash
			  /swflash.cab#version=6,0,29,0" width="34" height="18">
                          <param name="movie" value="swf/new.swf" />
                          <param name="quality" value="high" />
                          <param name="wmode" value="transparent" />
                          <embed src="swf/new.swf" width="34" height="18" quality="high"
				 pluginspage="http://www.macromedia.com/go/getflashplayer"
				  type="application/x-shockwave-flash" wmode="transparent"></embed>
                        </object>
                        <% else %>
                        <img src="images/send/read.gif" />
                        <% end if %>
                    </div></td>
                    <td>&nbsp;<a href="#" class="STYLE1" 
		  onclick="JScript:window.open('EveryDay/send_detail.asp?ID=<%=rs_email("ID")
		  %>','','width=560,height=400')"><%=rs_email("subject")%></a></td>
                    <td class="STYLE1"><div align="center" class="style2"><%=rs_email("SPerson")%></div></td>
                    <td class="STYLE1"><div align="center" class="STYLE1"><%=rs_email("sTime")%> </div></td>
                    <td><div align="center"><a href="#" 
		  onclick="JScript:window.open('EveryDay/send_del.asp?ID=<%=rs_email("ID")
		   %>','','width=560,height=400')"><img src="images/del.gif" width="16"
		    height="16" border="0" /></a></div></td>
                  </tr>
                  <% 
				  rs_email.movenext
				if rs_email.eof then exit for 
				next
				%>
                </table>
              <table width="100%" border="0" cellspacing="-2" cellpadding="-2">
                  <tr>
                    <td><div align="right" class="STYLE1">
                        <% if page1<>1 then %>
                      <a href="<%=path1%>?page1=1" class="STYLE1">第一页</a> <a href="<%=path1%>?page1=<%=(page1-1)%>" class="STYLE1">上一页</a>
                        <%end if 
		if page1<>rs_email.pagecount then %>
                        <a href="<%=path1%>?page1=<%=(page1+1)%>" class="STYLE1">下一页</a> <a href="<%=path1%>?page1=<%=rs_email.pagecount%>" class="STYLE1">最后一页</a>
                        <%end if %>
                      &nbsp;</div></td>
                  </tr>
                </table>
              <% end if %>
            </td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>