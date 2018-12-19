<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin:0px;
	background-color:#b8dbf9;
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
window.onload = function(){
	document.getElementById("message").focus();

}
</script>
<body  oncontextmenu=self.event.returnValue=false>
<table width="100%" height="115"  border="0" cellpadding="0" cellspacing="0">
<form action="main.asp" method="post" name="form1" target="mainFrame">
	<tr>
      <td width="75%" height="35">&nbsp;[<%=session("UserName")%>]
          <select name="face" >
            <option  value="无表情的">无表情的</option>
            <option value="微笑着" selected>微笑着</option>
            <option value="笑呵呵地">笑呵呵地</option>
            <option value="热情的">热情的</option>
            <option value="温柔的">温柔的</option>
            <option value="红着脸">红着脸</option>
            <option value="幸福的">幸福的</option>
            <option value="嘟着嘴">嘟着嘴</option>
            <option value="热泪盈眶的">热泪盈眶的</option>
            <option value="依依不舍的">依依不舍的</option>
            <option value="得意的">得意的</option>
            <option value="神秘兮兮的">神秘兮兮的</option>
            <option value="恶狠狠的">恶狠狠的</option>
            <option value="大声的">大声的</option>
            <option value="生气的">生气的</option>
            <option value="幸灾乐祸的">幸灾乐祸的</option>
            <option value="同情的">同情的</option>
            <option value="遗憾的">遗憾的</option>
            <option value="正义凛然的">正义凛然的</option>
            <option value="严肃的">严肃的</option>
            <option value="慢条斯理的">慢条斯理的</option>
            <option value="无精打采的">无精打采的</option>
          </select>
      对
	  <%
		  if Request.QueryString("name") <>"" then
			toname=Request.QueryString("name")
		  else 
			  toname="所有人"
		  end if
	  %>
      <input name="to" type="text" class="Sytle_text" id="to" value="<%=toname%>" size="20" readonly="readonly" style=" border-color: 8e999f">
      说      
      <input name="from" type="hidden" id="from" value="<%=session("UserName")%>">
	  </td>
      <td width="25%">&nbsp;字体颜色：
          <select name="color" size="1" id="select">
            <option selected>默认颜色</option>
            <option style="color:#FF0000" value="FF0000">红色热情</option>
            <option style="color:#0000FF" value="0000ff">蓝色开朗</option>
            <option style="color:#ff00ff" value="ff00ff">桃色浪漫</option>
            <option style="color:#009900" value="009900">绿色青春</option>
            <option style="color:#009999" value="009999">青色清爽</option>
            <option style="color:#990099" value="990099">紫色拘谨</option>
            <option style="color:#990000" value="990000">暗夜兴奋</option>
            <option style="color:#000099" value="000099">深蓝忧郁</option>
            <option style="color:#999900" value="999900">卡其制服</option>
            <option style="color:#ff9900" value="ff9900">镏金岁月</option>
            <option style="color:#0099ff" value="0099ff">湖波荡漾</option>
            <option style="color:#9900ff" value="9900ff">发亮蓝紫</option>
            <option style="color:#ff0099" value="ff0099">爱的暗示</option>
            <option style="color:#006600" value="006600">墨绿深沉</option>
            <option style="color:#999999" value="999999">烟雨蒙蒙</option>
        </select></td>
    </tr>
	<script language="javascript">
	function check(){
	if(form1.message.value=="")
	{alert("发言内容不能为空！");form1.message.focus();return;}
	form1.submit();  //提交表单
	form1.reset();  //重置表单
	form1.message.focus();
	}
	</script>
<script language="javascript">
	function Exit1(){
	parent.location.href="exit.asp";
	}
	</script>
    <tr>
      <td height="35">&nbsp;
          <input name="message" type="text" class="Sytle_text" id="message" onKeyDown="Javascript:if(event.keyCode==13){check()}" size="60" style="border-color: #8e999f;">
          <input name="send" type="button" id="send" value="&nbsp;" onClick="check()" style="background-image:url(images/send.gif); background-position:center; background-repeat:no-repeat; width:42px; height:21px; border: 1px #C67131 solid;"></td>
      <td><div align="center">
          <input type="button" name="Submit2" value="&nbsp;" onClick="Exit1()" style=" background-image:url(images/quit.gif); background-position:center; background-repeat:no-repeat; width:104px; height:23px; border: 1px #C67131 solid;">
      </div></td>
    </tr>
</form>	
    <tr>
      <td height="35" colspan="2">
        <table>
	  <tr>
	  <td>
<form action="Upload_OK.asp" method="post" enctype="multipart/form-data" name="form2" target="mainFrame">
	  请选择图片:
	  <input name="filename" type="file" id="filename" size="40" class="Sytle_text" style=" border-color: #8e999f;">
	  <input type="button" name="SendImg" value="&nbsp;" onClick="JavaScript:form2.submit();" style=" background-image:url(images/sendpic.gif); background-position:center; background-repeat:no-repeat; width:61px; height:21px; border: 1px #C67131 solid;">
</form>	  
	  </td>
	  </tr>
	  </table>
	</td>
    </tr>
  </table>
  </body>