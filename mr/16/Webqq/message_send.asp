<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql="SELECT * from tb_user where ID="&request.QueryString("UserID")&""
rs.open sql,conn,1,3
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<script language="javascript">
	window.resizeTo("410","211")					//控制弹出页的大小
	estate = "close"								//为estate变量赋值默认值
        function talknote()
		{//创建自定义函数									
	     if(estate=="close")						//判断estate变量的值是否为“close”
		 {
			window.resizeTo("410","411")			//控制弹出页的大小
			document.all.dibu.style.display=""	//显示聊天记录页面
			window.frames("bottom").location = "talknote.asp?UserID=<%=request.QueryString("UserID")%>"
			//重新加载聊天记录页面
			estate = "open"							//为estate变量赋默认值
	     }
	     else{
			window.resizeTo("410","211")
			document.all.dibu.style.display="none"  //隐藏聊天记录页面
			estate = "close"						//为estate变量赋默认值
	     }
	}
	function mycheck1(){
	}
	
	
	
	function mycheck2()
	{
	   if(trim(document.form1.message.value)==""){
	      alert("不可以发送空信息!")
		  return;
	   }
	   else{
	     document.form1.submit()
	   }
	}
	function formSubmit()
	{									//创建自定义函数
	    keyvalue=window.event.keyCode	
		if((keyvalue==13) && window.event.ctrlKey)
		{
		  	   if(trim(document.form1.message.value)=="")		//获取信息文本框的值是否为空
			  {
				  alert("不可以发送空信息!")					//弹出提示对话框
				  document.form1.message.focus()				//获取焦点
				  return;
			   }
			   else
			   {
			     window.moveTo(2000,200) 						//移动窗口
				 document.form1.submit()						//提交表单
                 document.all.btnsubmit.disabled=true;			//设置发送信息按钮为不可用的状态
			   }
		}
		else{
			return
		}
	}
function  trim(strInput){ 
	var iLoop=0;
	var iLoop2=-1;
	var strChr;
	if((strInput == null)||(strInput == "<NULL>"))
		return "";
	if(strInput){
		for(iLoop=0;iLoop<strInput.length-1;iLoop++){
			strChr=strInput.charAt(iLoop);
			if(strChr!=' ')
				break;
		}
		for(iLoop2=strInput.length-1;iLoop2>=0;iLoop2--){
			strChr=strInput.charAt(iLoop2);
			if(strChr!=' ')
				break;
		}
	}
	if(iLoop<=iLoop2){
		return strInput.substring(iLoop,iLoop2+1);
	}
	else{
		return "";
	}
}
function closed(){
 window.moveTo(2000,200)
}
</script>
<style type="text/css">
<!--
.Fieldset {  font-size: 12px; border: 1px inset; padding-top: 2px; padding-right: 4px; padding-bottom: 2px; padding-left: 4px; margin-top: 2px; margin-right: 4px; margin-bottom: 2px; margin-left: 4px}
.button {  font-size: 12px; text-align: center; vertical-align: middle}
td {  padding-right: 2px; padding-left: 2px}
.text {  background-color: #ECE9D8; height: 18px}
.STYLE1 {
	font-size: 9pt;
	color: #000000;
}
.STYLE2 {font-size: 9pt}
-->
</style>
</head>
<body scroll="no" topmargin="0" leftmargin="0" style="border:none" bgcolor="#ECE9D8" id="bodyID" oncontextmenu="mycheck1()" onkeydown="formSubmit()">
<table border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" width="100%" id="AutoNumber1" height="100%" style="border-collapse: collapse">
  <form name="form1" action="messageSend_deal.asp" method="post">
  <tr height="50">
    <td bgcolor="BBD2FC">
     <fieldset class="Fieldset" style="border: 1px# #999999 solid;"><legend><span lang="zh-cn">发送给</span></legend> 
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr> 
          <td width="25%" bgcolor="BBD2FC">
            <div align="right" class="STYLE1">用户名:</div>          </td>
          <td width="25%" bgcolor="BBD2FC" class="STYLE1">
            <a href="#" onClick="javascript:window.open('chakan.asp?id=<%=rs("ID")%>&type=new','','width=350,height=250,toolbar=no,location=no,status=no,menubar=no')"><%=rs("UserName")%></a>
          </td>
          <td width="25%" bgcolor="BBD2FC">
            <div align="right" class="STYLE1 STYLE2">昵称:</div>          </td>
          <td width="25%" bgcolor="BBD2FC">
            <span class="STYLE2"><%=rs("state")%>            </span>
            <input type="hidden" value="<%=rs("ID")%>" name="ToUserID">
			<input type="hidden" value="<%=Session("username1")%>" name="fusernickname">
			<input type="hidden" value="<%=Session("user_id")%>" name="fuserid"></td>
        </tr>
      </table>
      </fieldset>
    </td>
  </tr>
  <tr id="topTr">
    <td width="100%"><textarea name="message" style="width:100%;height:100%" onKeyDown="document.all.btnsubmit.disabled=false;"></textarea></td>
  </tr>
  <tr height="28" id="barTr">
    <td width="100%">
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
          <td bgcolor="BBD2FC">
            <input type="button" value="聊天记录" name="B3" onClick="talknote()" class="button"></td>
			<td align="right" bgcolor="BBD2FC">
			&nbsp;&nbsp;&nbsp;
			<input type="button" value="发送消息" name="B3" class="button" onClick="mycheck2();this.disabled=true;" id="btnsubmit">
			<input type="button" value="关闭" name="B3" class="button" onClick="closed()">          </td>
		</tr>
    </table>
    </td>
  </tr>
  <tr id="dibu" height="200" style="display:none">
    <td width="100%"><iframe name="bottom" style="width:100%;height:100%" src=""></iframe></td>
  </tr>
  </form>
</table>
</body>
<script>
document.form1.message.focus()
</script>
</html>