<html xmlns:sml>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Htc/NavBar.css">
<title>������Ⱥ����</title>
<script>
function fnAddGroup(){
	Url = "webQQ_AddGroup.asp"
	Name= "addGroup"
	features = "width=270,height=100"
	window.open(Url,Name,features) 
}

function fnAddGroupMember(){
    if(!window.frames("mainfrm").frames("right").document.form1){
    alert("��ѡһ����Ҫ��ӵ�Ŀ����!")
    return
    }
	groupID = window.frames("mainfrm").frames("right").document.form1.groupID.value
	Url = "webQQ_AddGroupMember.asp?GroupID=" + groupID
	Name= "addGroupMember"
	features = "width=270,height=100"
	window.open(Url,Name,features) 
}

function fndelGroup(){
if(!window.frames("mainfrm").frames("left").document.form1.groupID){
	return;
}
    obj = window.frames("mainfrm").frames("left").document.form1.groupID
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
	 alert("������ѡ��һ������ɾ��!")
	 return
	 }

	window.frames("mainfrm").frames("left").document.form1.submit()
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
</head>
<body topmargin="0" leftmargin="0" style="border:none">

<table border="0" cellpadding="0" cellspacing="0" width="100%" id="AutoNumber1" height="100%">
<tr height="24" id="NavbarTr">
  <td class="indextdTop">
    <sml:Menubar>
     <Root>
       <MenuBar url="javascript:fnAddGroup()" img="addgroup.gif" name="�����"></MenuBar>
       <MenuBar url="javascript:fndelGroup()" img="Deletegroup.gif" name="ɾ����"></MenuBar>
       <MenuBar url="javascript:fnAddGroupMember()" img="AddUser.gif" name="��Ӻ���"></MenuBar>
       <MenuBar url="javascript:fndelMem()" img="Deleteuser.gif" name="ɾ������"></MenuBar>
       <MenuBar url="javascript:window.close()" img="Exit.gif" name="�˳�"></MenuBar>
     </Root>
     </sml:Menubar>
  </td>
  </tr>
  <tr>
    <td width="100%"><iframe name="mainfrm" width="100%" height="100%" frameborder="0" src="webQQ_GroupFrame.asp"></iframe></td>
  </tr>
</table>

</body>

</html>