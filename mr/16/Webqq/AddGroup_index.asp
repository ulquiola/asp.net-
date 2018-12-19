<%@ Language=VBScript %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>添加群组</title>
<script>
function mychecksubmit(){
	if(trim(document.form1.GroupName.value)==""){
	alert("请输入分组名称!")
	return;
	}
	document.form1.submit()
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
</script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>
<style type="text/css">
<!--
.STYLE2 {
	font-size: 9pt;
	color: #000000;
}
-->
</style>
<body  style="BORDER-RIGHT:  medium none; BORDER-TOP: medium none; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none">
<table width="272" height="109" border="0" cellpadding="0" cellspacing="0" background="Images/addgroup.gif">
  <tr>
    <td height="109"><form name=form1 action="AddGroup_deal.asp" method="post">
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%"><br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="GroupName"></td>
  </tr>
</table>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="button" value="提交" name="B1" onClick="mychecksubmit()">
</form></td>
  </tr>
</table>
</body>

</html>