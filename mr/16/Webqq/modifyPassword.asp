<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>密码修改</title>
<script>
window.resizeTo("270","180")
function modifypassword(){
   if(form1.password.value==""||form1.password1.value==""){
      alert("密码和确认密码不能为空!")
      return
   }
   if(form1.password.value!=form1.password1.value){
      alert("两次输入的密码不一致，请重新输入!")
	  return
   }
   else{
      document.form1.action = "modifypassword_deal.asp"
	  document.form1.submit()
   }
}
</script>
<style type="text/css">
<!--
.STYLE2 {font-size: 9pt}
.STYLE4 {font-size: 9pt; color: #000000; }
-->
</style>
</head>
<body topmargin="0" leftmargin="0" style="border:none" scroll=no>
<form name="form1" action="" method="POST">
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber1" height="70%">
    <tr> 
      <td width="25%" height="16" align="right"><span class="STYLE2">新密码:</span></td>
      <td width="75%" align="left"><font size="2"> 
        <input name="password" type="password" id="password" style="width:95%">
        </font></td>
    </tr>
    <tr> 
      <td width="25%" height="24" align="right"><span class="STYLE4">确认密码:</span></td>
      <td width="75%" align="left"><font size="2"> 
        <input name="password1" type="password" id="password1" style="width:95%">
        </font></td>
    </tr>
    <tr> 
      <td height="45" colspan="2" align="right">
        <div align="center">
          <input type="button" value="提交" name="B3" onClick="modifypassword()">
          &nbsp;
          <input type="button" value="取消" name="B3" onClick="window.close()">
      &nbsp;&nbsp; </div></td></tr>
  </table>
</form>
</body>

</html>