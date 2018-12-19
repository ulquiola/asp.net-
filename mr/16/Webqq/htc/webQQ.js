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



function fnOpenWin(UserID)
{
if(readmessage.closed){
        readmessage=window.open("blank.asp","readmessage","width=410,height=211")
		readmessage.moveTo(2000,400)

}
  objImg = eval("document.all.Img"+UserID)
  alert(objImg.src.indexOf("_00.gif")!=-1);
  if(objImg.src.indexOf("_00.gif")!=-1){
	  theURL = "readmessage.asp?UserID="+UserID;
	   readmessage.location = theURL
	  readmessage.moveTo(200,200)
  }
  else{
	  //theURL = "message_send.asp?UserID="+UserID;
	 //  readmessage.location = theURL
	//  readmessage.moveTo(200,200)
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