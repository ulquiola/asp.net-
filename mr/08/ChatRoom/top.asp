
<style type="text/css">
<!--
body {
	margin:0px;
}
-->
</style>
<script language="javascript">
function enterkey(){
	if(event.keyCode == 116){				//如果按键是〈F5〉键
		alert('禁止刷新');				//弹出警告框
		event.keyCode = 0;				//将按键归零
		return false;
	}
}
document.onkeydown=enterkey;				//将函数赋给onkeydown事件
</script>
<body oncontextmenu=self.event.returnValue=false bgcolor="#000000">
<div style=" background-image:url(images/header.jpg); background-position: top; background-repeat:no-repeat; width:785px; height:133px;"></div>
</body>