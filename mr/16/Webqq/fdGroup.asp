<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html xmlns:sml>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ѷ������</title>
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
function mycheckAdd(){				//����mycheckadd�Զ��庯��
	Url = "AddGroup_index.asp"		//ҳ���URL
	Name= "addGroup"				//ҳ�������
	size = "width=270,height=100"	//ҳ��Ĵ�С
	window.open(Url,Name,size) 		//��һ��ָ����С��ASPҳ��
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
	 alert("��ѡһ����Ҫ��ӵ�Ŀ����!")
	 return
	 }
 
	groupID = window.frames("mainfrm").frames("right").document.form1.groupID.value
	Url = "addfriend_index.asp?GroupID=" + groupID
	Name= "index"
	size = "width=470,height=320"
	window.open(Url,Name,size) 
}



function mycheckdel()
{//�����Զ��庯��
if(!window.frames("mainfrm").frames("left").document.form1.groupID){
	return;
}
    oo=window.frames("mainfrm").frames("left").document.form1.groupID
   //��ȡ���ID��
    var Ced = "0"
	//ced������ֵ
	if(oo.length)
	{
	 for(i=0;i<oo.length;i++){
	  if(oo[i].checked){
	   Ced = "1"
	   //ced��ֵΪ��1��ʱ��ʾ��ѡ����ѡ��
	  }
	 }
	}
	else{
	 if(oo.checked){
	  Ced = "1"
	 }
	 }
	 if(Ced!="1"){
	//ced��ֵ��Ϊ��1��ʱ��ʾ��ѡ��δѡ��
	 alert("��ѡ��һ�����ٽ���ɾ��!")
	 //������ʾ�Ի���
	 return
	 }
	window.frames("mainfrm").frames("left").document.form1.submit()
	//�ύ��
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
	 alert("������ѡ��һ��������ɾ��!")
	 return
	 }
	window.frames("mainfrm").frames("right").dodelete()
}
</script>
<body topmargin="0" leftmargin="0" style="border:none">
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="AutoNumber1" height="100%">
<tr height="24" id="NavbarTr">
  <td width="16%" height="22" background="Images/bgfg.gif">
    <img src="Images/group1.gif" width="26" height="21">&nbsp;<a href="#" class="STYLE1" onClick="javascript:mycheckAdd()">�����</a></td>
  <td width="17%" background="Images/bgfg.gif"><img src="Images/group2.gif" width="26" height="22">&nbsp;<a href="#" class="STYLE1" onClick="javascript:mycheckdel()">ɾ����</a></td>
  <td width="19%" background="Images/bgfg.gif"><img src="Images/group3.gif" width="26" height="22">&nbsp;<a href="#" class="STYLE1" onClick="javascript:addfriend()">��Ӻ���&nbsp;</a></td>
  <td width="21%" background="Images/bgfg.gif"><img src="Images/group4.gif" width="23" height="22">&nbsp;<a href="#" class="STYLE1" onClick="javascript:fndelMem()">ɾ������&nbsp;</a></td>
  <td width="15%" background="Images/bgfg.gif"><img src="Images/group5.gif" width="22" height="22">&nbsp;<a href="#" class="STYLE1" onClick="javascript:window.close()">�˳�&nbsp;</a></td>
  <td width="12%" background="Images/bgfg.gif" class="indextdTop">&nbsp;</td>
</tr>
  <tr>
    <td colspan="6"><iframe name="mainfrm" width="100%" height="100%" frameborder="0" src="GFrame.asp"></iframe></td>
  </tr>
</table>

</body>
</html>
