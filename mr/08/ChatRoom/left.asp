<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin:12px;
	font-size:12px;
	background-color: #f0f5fb;
}
-->
</style>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<script type="text/javascript" src="js/AjaxRequest.js"></script>
<script type="text/javascript">
timer = window.setInterval("showLeftAJAX()",2000);
function showLeftAJAX(){
	var left_list = new net.AjaxRequest("show.asp?random="+Math.random(),showLeftAll,"GET","");
}
function showLeftAll(){
	document.getElementById("show").innerHTML = this.req.responseText;
}

function enterkey(){
	if(event.keyCode == 116){				//如果按键是〈F5〉键
		alert('禁止刷新');					//弹出警告框
		event.keyCode = 0;					//将按键归零
		return false;
	}
}
function chkname(fname,sname){
	if(fname == sname){
		alert("不允许自言自语");
		return false;
	}
	var sendimg = new net.AjaxRequest("UpLoad_OK1.asp?random="+Math.random()+"&Toimg="+sname,addobj,"GET","");	
}
function addobj(){
}

document.onkeydown=enterkey;

</script>
<body  oncontextmenu=self.event.returnValue=false>
<div style=" padding:12px;">
欢迎您走进聊天室！<br/>
</div>
<div id="show">
<%
A_listName=split(Application("UserName"),",")
A_head=split(Application("head"),",")
%>
<div style="text-align:left; padding:12px;">目前共有<font color="#CC0000"><b><%=UBound(A_ListName)%></b></font>人在线</div>
<%
For i=0 To UBound(A_ListName)%>
<div style="text-align:left; text-indent: 25px;">
	<a href="bottom.asp?name=<%=A_ListName(i)%>" target="bottomFrame" onClick="return chkname('<%=session("UserName")%>','<%=A_ListName(i)%>')">
	<%HeadPath="headICO/"&A_head(i)%>
	<img src="<%=HeadPath%>" border="0" width="23" height="23">
	<%=A_ListName(i)%></a>
</div>
<%Next%>
</div>
<br>
</body>