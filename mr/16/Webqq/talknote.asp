<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
       <%
		Set rssml = Server.CreateObject("ADODB.Recordset")		'创建记录集
		rssml.ActiveConnection = conn							'定义Command对象的连接信息
		rssml.Source = "SELECT * from tb_talk where (fuserid = "&request.QueryString("UserID")&" and ToUserID = " &request.Cookies("UserID")&") or (fuserid = "&request.Cookies("UserID")&" and ToUserID = " &request.QueryString("UserID")&") order by id desc"
		'查询数据库信息
		rssml.CursorType = 0									'定义游标类型
		rssml.CursorLocation = 2								'定义游标访问区域
		rssml.LockType = 1										'RecordSet对象为只读，不允许作任何修改
		rssml.Open()											'打开记录集
		%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script>
function clearmessage()
{
	window.frames("frmrefrash").location = "clearmessage.asp?UserID=<%=request.QueryString("UserID")%>"
}
</script>
<style type="text/css">
<!--
.STYLE2 {
	font-size: 9pt;
	color: #000000;
}
-->
</style>
</head>

<body topmargin="0">
<div id="Layer1" style="position:absolute; width:200px; height:115px; z-index:1; visibility: hidden;">
<iframe name="frmrefrash" src=""></iframe>
</div>
<span class="STYLE2">[<a href="javascript:clearmessage()">清除所有聊天记录</a>]</span> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE2">[ <a href="#" onClick="window.clipboardData.setData('Text',document.getElementById('content').value);">复制当前聊天记录</a> ]</span><br>
<span class="STYLE2">
<%
while not rssml.EOF								'查询是否有记录信息
mes=rssml.Fields.Item("Message").value			'获取字段信息
%>	
<font color="#FF0000"><%=rssml("SendDate")%>&nbsp;<%=rssml("fusernickname")%> </font><br>
<%'应用特殊的颜色显示发送信息日期和用户昵称%>
<font color="#3333FF"><%=mes%> </font><br>
<%'应用特殊的颜色显示聊天记录信息%>
<%
content=content&rssml("SendDate")&"&nbsp;&nbsp;"&rssml("fusernickname")&chr(13)&mes&chr(13)&chr(13)
'为content变量赋值
rssml.movenext()
'向下移动记录指针
wend
rssml.close()
'关闭记录集
set rssml=nothing
'释放内存空间
%>
</span>
<input type="hidden" name="content" id="content" value="<%=content%>">
</body>
</html>
