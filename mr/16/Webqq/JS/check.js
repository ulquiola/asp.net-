//�ж��û���¼��Ϣ�Ƿ�Ϊ��
function mycheck1()
{
   if (form1.UserName.value==""){
		alert("�������û�����");form1.UserName.focus();return;
	}
	if (form1.Password1.value==""){
		alert("���벻��Ϊ��!");form1.Password1.focus();return;
	}
   else{
      document.form1.action = "User_cl.asp"
	  document.form1.submit()
   }
}
//�ж�����������Ƿ���ȷ
function CheckDate(INDate)
{ if (INDate=="")
    {return true;}
	subYY=INDate.substr(0,4)
	if(isNaN(subYY) || subYY<=0){
		return true;
	}
	//ת���·�
	if(INDate.indexOf('-',0)!=-1){	separate="-"}
	else{
		if(INDate.indexOf('/',0)!=-1){separate="/"}
		else {return true;}
		}
		area=INDate.indexOf(separate,0)
		subMM=INDate.substr(area+1,INDate.indexOf(separate,area+1)-(area+1))
		if(isNaN(subMM) || subMM<=0){
		return true;
	}
		if(subMM.length<2){subMM="0"+subMM}
	//ת����
	area=INDate.lastIndexOf(separate)
	subDD=INDate.substr(area+1,INDate.length-area-1)
	if(isNaN(subDD) || subDD<=0){
		return true;
	}
	if(eval(subDD)<10){subDD="0"+eval(subDD)}
	NewDate=subYY+"-"+subMM+"-"+subDD
	if(NewDate.length!=10){return true;}
    if(NewDate.substr(4,1)!="-"){return true;}
    if(NewDate.substr(7,1)!="-"){return true;}
	var MM=NewDate.substr(5,2);
	var DD=NewDate.substr(8,2);
	if((subYY%4==0 && subYY%100!=0)||subYY%400==0){ //�ж��Ƿ�Ϊ����
		if(parseInt(MM)==2){
			if(DD>29){return true;}
		}
	}else{
		if(parseInt(MM)==2){
			if(DD>28){return true;}
		}	
	}
	var mm=new Array(1,3,5,7,8,10,12); //�ж�ÿ���е��������
	for(i=0;i< mm.length;i++){
		if (parseInt(MM) == mm[i]){
			if(parseInt(DD)>31){return true;}
		}else{
			if(parseInt(DD)>30){return true;}
		}
	}
	if(parseInt(MM)>12){return true;}
   return false;
}
//�ж������Email��ַ��ʽ�Ƿ���ȷ
function checkemail(Email)
{
	var str=Email;
	 //��JavaScript�У�������ʽֻ��ʹ��"/"��ͷ�ͽ���������ʹ��˫����
	var Expression=/\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*/; 
	var objExp=new RegExp(Expression);
	if(objExp.test(str)==true){
		return true;
	}else{
		return false;
	}
}	
//�ж��û�ע����Ϣ�Ƿ�Ϊ��
 function mycheck(){
		  	if(myform.UserName.value==""){
				alert("�������û�����");myform.UserName.focus();return;	}
			if(myform.Password1.value==""){
				alert("���������룡");myform.Password1.focus();return;}
		  	if(myform.Password2.value==""){
				alert("������ȷ�����룡");myform.Password2.focus();return;}
		  	if(myform.Password1.value!=myform.Password2.value){
				alert("��������������벻һ�£����������룡");
				myform.Password1.value="";myform.Password2.value="";myform.Password1.focus();return;}
			if(myform.birthday.value==""){
				alert("��������������!");myform.birthday.focus();return;
			}		
			if(CheckDate(myform.birthday.value)){
				alert("����������ڸ�ʽ����ȷ���磺1980/03/13��1980-03-13��\n ����������!");myform.birthday.focus();return;
			}
			if(myform.Email.value==""){
				alert("������Email��ַ��");myform.Email.focus();return;	}
			if(!checkemail(myform.Email.value)){
				alert("�������Email��ַ����ȷ������������!");myform.Email.focus();return;
			}
			myform.submit();
}

var FrameTrindex = null;
function moveup(){
   if(FrameTrindex==null){
     objMask = mask;
     objMenu = menu;
     objContact = contact;
   }
   else{
     objMask = mask[FrameTrindex];
     objMenu = menu[FrameTrindex];
     objContact = contact[FrameTrindex];
   }
	if(objMask&&objMenu){
		 menuHeight = objMask.offsetHeight
		 actualHeight=objMenu.offsetHeight
		 if(document.all&&objMenu.style.pixelTop>(menuHeight-actualHeight)){
		   objMenu.style.pixelTop-=5
		 }
		 else{
		   clearInterval(objContact.doing)
		 } 
	}
}

function movedown(){
   if(FrameTrindex==null){
     objMask = mask;
     objMenu = menu;
     objContact = contact;
   }
   else{
     objMask = mask[FrameTrindex];
     objMenu = menu[FrameTrindex];
     objContact = contact[FrameTrindex]
   }
	if(objMask&&objMenu){
		menuHeight = objMask.offsetHeight
		actualHeight=objMenu.offsetHeight
		if(document.all&&objMenu.style.pixelTop<0){
		  objMenu.style.pixelTop+=5
		}
		else{
		  clearInterval(objContact.doing)
		}
	}
}

function doMoveUp(obj){
    obj = window.event.srcElement.parentElement
    if(document.all.upDiv.length){
		for(i=0;i<document.all.upDiv.length;i++){
		   if(obj==document.all.downDiv[i]){
		       FrameTrindex = i 
		   }
		} 
		contact[FrameTrindex].doing = setInterval("moveup()",10)
    }
    else{
        contact.doing = setInterval("moveup()",10)
    }
    
    
}

function doMoveDown(obj){
obj = window.event.srcElement.parentElement
  if(document.all.upDiv.length){
    for(i=0;i<document.all.upDiv.length;i++){
       if(obj==document.all.upDiv[i]){
           FrameTrindex = i 
       }
    }
    contact[FrameTrindex].doing = setInterval("movedown()",10)
  }
  else{
    contact.doing = setInterval("movedown()",10)
  }
}

function doStopMove(){
   if(FrameTrindex==null){
     objContact = contact;
   }
   else{
     objContact = contact[FrameTrindex];
   }
   if(objContact.doing){
      clearInterval(objContact.doing)
   }
}

function hideAll(){
  for(i=0;i<document.all.FrameTr.length;i++){
     document.all.FrameTr[i].style.display="none"
  }
}

function showSelectedItem(Item){
    if(!document.all.FrameTr.length){return;}
    hideAll()
    eval('document.all.FrameTr['+Item+'].style.display=""')
    initScrollBar(Item)
}

function initScrollBar(Item){
         menuHeight = mask[Item].offsetHeight
		 actualHeight=menu[Item].offsetHeight
		 if(actualHeight>menuHeight){
		   scrollbar[Item].style.display=""
		 }
		 else{
		   scrollbar[Item].style.display="none"
		 }
}

function fnOpenWin(UserID){
if(readmessage.closed){
        readmessage=window.open("blank.asp","readmessage","width=410,height=211")
		readmessage.moveTo(2000,400)

}
  objImg = eval("document.all.Img"+UserID)
  if(objImg.src.indexOf("_00.gif")!=-1){
	  tempStr1=objImg.src.substring(0,objImg.src.lastIndexOf("_"))+".gif";
	  objImg.src=tempStr1;
	  theURL = "readmessage.asp?UserID="+UserID;
	   readmessage.location = theURL
	  readmessage.moveTo(200,200)
  }
  else{
	  theURL = "message_send.asp?UserID="+UserID;
	   readmessage.location = theURL
	  readmessage.moveTo(200,200)
  }
}

function unline(obj){
   UserID = obj.UserID
   HasNewMsg = obj.HasNewMsg
   if(HasNewMsg=="0"){
   obj.UserID = ""
	return;
   }
   document.all.unknowMsg.style.display = "none"
   obj.HasNewMsg = "0"
   theURL = "readmessage.asp?UserID="+UserID;
    readmessage.location = theURL
    readmessage.moveTo(200,200)
}

function biancss(obj){
  obj.orgclassName = obj.className
  obj.className = "ItemHL"
}

function huifuCss(obj){
   obj.className = obj.orgclassName
}

function resetFace(UserID){
    objImg = eval("document.all.Img"+UserID)
    objImg.src=objImg.src
}