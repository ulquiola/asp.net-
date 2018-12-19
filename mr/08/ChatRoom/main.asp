<%
'处理用户发送信息的代码
If request.Form("from")<>"" Then
	formname=Request.Form("from")
	toName=Request.Form("To")
	face=Request.Form("face")
	color=Request.Form("color")
	message=server.HTMLEncode(Request.Form("message"))
	content=formname&" <b><font color='#0000FF'>"&face&_
	"</font></b> 对[<font color='#FF0000'>"&toName&"</font>] 说：<font color=#"&_
	color&">"&message&"</font> ("&now&")"
	If application("message")="" Then
		application("message")=content
	Else
		application("message")=application("message")+"|"+content
	End If
End If
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	background-color: #ffffff;
}
.style1 {color: #FF0000}
.style2 {
	color: #0000FF;
	font-style: italic;
}
-->
</style>
<script type="text/javascript" src="JS/AjaxRequest.js"></script>
<script type="text/javascript">
timer = window.setInterval("showChatAJAX()",1000); 
function showChatAJAX(){
	var main_list = new net.AjaxRequest("showchat.asp?random="+Math.random(),showChatAll,"GET","");
}
function showChatAll(){
	document.getElementById("chat_show").innerHTML = this.req.responseText;
}
function enterkey(){
		if(event.keyCode == 116){				//如果按键是〈F5〉键
			alert('禁止刷新');				//弹出警告框
			event.keyCode = 0;				//将按键归零
			return false;
		}
	}
document.onkeydown=enterkey;
</script>
</head>
<body  oncontextmenu=self.event.returnValue=false>
<font color="#CC0000" size="2">为了聊天室的正常运行!请聊天者正常退出!除非遇到紧急情况!谢谢合作!!</font>
<br>
<table>
 <tr>
  <td>
  <div id="chat_show">
<%
A_message=split(Application("Message"),"|")
for i=uBound(A_message) to StartV step -1
	response.Write(A_message(i)+"<BR>")
Next
%>
</div>
  </td>
  </tr>
</table>
</body>