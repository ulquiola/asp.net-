<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�����޸�</title>
<script>
window.resizeTo("270","180")
function modifypassword(){
   if(form1.password.value==""||form1.password1.value==""){
      alert("�����ȷ�����벻��Ϊ��!")
      return
   }
   if(form1.password.value!=form1.password1.value){
      alert("������������벻һ�£�����������!")
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
      <td width="25%" height="16" align="right"><span class="STYLE2">������:</span></td>
      <td width="75%" align="left"><font size="2"> 
        <input name="password" type="password" id="password" style="width:95%">
        </font></td>
    </tr>
    <tr> 
      <td width="25%" height="24" align="right"><span class="STYLE4">ȷ������:</span></td>
      <td width="75%" align="left"><font size="2"> 
        <input name="password1" type="password" id="password1" style="width:95%">
        </font></td>
    </tr>
    <tr> 
      <td height="45" colspan="2" align="right">
        <div align="center">
          <input type="button" value="�ύ" name="B3" onClick="modifypassword()">
          &nbsp;
          <input type="button" value="ȡ��" name="B3" onClick="window.close()">
      &nbsp;&nbsp; </div></td></tr>
  </table>
</form>
</body>

</html>