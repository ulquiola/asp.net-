<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html xmlns:sml>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>好友分组管理</title>
<style type="text/css">
<!--
.STYLE1 {
	font-size: 9pt;
	color: #000000;
}
-->
</style>
</head>
<style type="text/css">
<!--
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
<script language="javascript">
function mycheckAdd(){				//创建mycheckadd自定义函数
	Url = "AddGroup_index.asp"		//页面的URL
	Name= "addGroup"				//页面的名称
	size = "width=270,height=100"	//页面的大小
	window.open(Url,Name,size) 		//打开一个指定大小的ASP页面
}

function addfriend()
{
  if(!window.frames("mainfrm").frames("right").document.form1){
   return;
  }
      oo = window.frames("mainfrm").frames("left").document.form1.groupID
    var Ced = "0"
	if(oo.length){
	 for(i=0;i<oo.length;i++){
	  if(oo[i].checked){
	   Ced = "1"
	  }
	 }
	}
	else{
	 if(oo.checked){
	  Ced = "1"
	 }
	 }
	 if(Ced!="1"){
	 alert("请选一个想要添加的目标组!")
	 return
	 }
 
	groupID = window.frames("mainfrm").frames("right").document.form1.groupID.value
	Url = "addfriend_index.asp?GroupID=" + groupID
	Name= "index"
	size = "width=470,height=320"
	window.open(Url,Name,size) 
}



function mycheckdel()
{//创建自定义函数
if(!window.frames("mainfrm").frames("left").document.form1.groupID){
	return;
}
    oo=window.frames("mainfrm").frames("left").document.form1.groupID
   //获取组的ID号
    var Ced = "0"
	//ced变量赋值
	if(oo.length)
	{
	 for(i=0;i<oo.length;i++){
	  if(oo[i].checked){
	   Ced = "1"
	   //ced的值为“1”时表示复选框已选中
	  }
	 }
	}
	else{
	 if(oo.checked){
	  Ced = "1"
	 }
	 }
	 if(Ced!="1"){
	//ced的值不为“1”时表示复选框未选中
	 alert("请选择一个组再进行删除!")
	 //弹出提示对话框
	 return
	 }
	window.frames("mainfrm").frames("left").document.form1.submit()
	//提交表单
}



function fndelMem(){
    if(!window.frames("mainfrm").frames("right").document.form1){
	return;
   }
   if(!window.frames("mainfrm").frames("right").document.form1.MemberID){
   return;
   }
    obj = window.frames("mainfrm").frames("right").document.form1.MemberID
    var Ced = "0"
	if(obj.length){
	 for(i=0;i<obj.length;i++){
	  if(obj[i].checked){
	   Ced = "1"
	  }
	 }
	}
	else{
	 if(obj.checked){
	  Ced = "1"
	 }
	 }
	 if(Ced!="1"){
	 alert("请至少选择一个好友再删除!")
	 return
	 }
	window.frames("mainfrm").frames("right").dodelete()
}
</script>
<body topmargin="0" leftmargin="0" style="border:none">
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="AutoNumber1" height="100%">
<tr height="24" id="NavbarTr">
  <td width="16%" height="22" background="Images/bgfg.gif">
    <img src="Images/group1.gif" width="26" height="21">&nbsp;<a href="#" class="STYLE1" onClick="javascript:mycheckAdd()">添加组</a></td>
  <td width="17%" background="Images/bgfg.gif"><img src="Images/group2.gif" width="26" height="22">&nbsp;<a href="#" class="STYLE1" onClick="javascript:mycheckdel()">删除组</a></td>
  <td width="19%" background="Images/bgfg.gif"><img src="Images/group3.gif" width="26" height="22">&nbsp;<a href="#" class="STYLE1" onClick="javascript:addfriend()">添加好友&nbsp;</a></td>
  <td width="21%" background="Images/bgfg.gif"><img src="Images/group4.gif" width="23" height="22">&nbsp;<a href="#" class="STYLE1" onClick="javascript:fndelMem()">删除好友&nbsp;</a></td>
  <td width="15%" background="Images/bgfg.gif"><img src="Images/group5.gif" width="22" height="22">&nbsp;<a href="#" class="STYLE1" onClick="javascript:window.close()">退出&nbsp;</a></td>
  <td width="12%" background="Images/bgfg.gif" class="indextdTop">&nbsp;</td>
</tr>
  <tr>
    <td colspan="6"><iframe name="mainfrm" width="100%" height="100%" frameborder="0" src="GFrame.asp"></iframe></td>
  </tr>
</table>

</body>
</html>
