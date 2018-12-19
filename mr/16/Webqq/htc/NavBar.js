function moveleft(mask,menu){
if(mask&&menu){
menuwidth = mask.offsetWidth
actualwidth=menu.offsetWidth
if (document.all&&menu.style.pixelLeft>(menuwidth-actualwidth))
menu.style.pixelLeft-=150
}
}

function moveright(mask,menu){
if(mask&&menu){
menuwidth = mask.offsetWidth
actualwidth=menu.offsetWidth
if(document.all&&menu.style.pixelLeft<0)
menu.style.pixelLeft+=150
}
}

orgButton = null
orgBar = null

function clickbutton(){
  //alert()
  ele = window.event.srcElement
   while(ele){
		if(ele.tagName=="TABLE"){				
					break;
		}
		ele=ele.parentElement;
	}

//==================================================
  CurrBtn = ele
  CurrBar = ele.parentElement
  
   if(CurrBtn.state=="off"){
    if(orgButton!=null){
      orgButton.state = "off"
      orgButton.className = "ButtonNM" 
      orgBar.className = "BarTDNM" 
   }
    orgButton = CurrBtn
    orgBar = CurrBar 
    
    CurrBtn.className="ButtonClk"
    CurrBar.className="BarTDClk"
    CurrBtn.state = "on"
  }
}


function biancss(){
    ele = window.event.srcElement
   while(ele){
		if(ele.tagName=="TABLE"){				
					break;
		}
		ele=ele.parentElement;
	}
   obj= ele
  if(obj.state=="off"){
  obj.orgclassName = obj.className
  obj.className = "ButtonHL"
  }
}
function huifuCss(){
    ele = window.event.srcElement
   while(ele){
		if(ele.tagName=="TABLE"){				
					break;
		}
		ele=ele.parentElement;
	}
   obj= ele
  if(obj.state=="off"){
  obj.className = obj.orgclassName
  }
}

function fnShow(obj){
  url = obj.Url
  if(url.indexOf("javascript:")==-1){
  window.frames("main").location.href = url + "?ID=" + document.form1.ID.value
  }
  else{
  jscmd = url.replace("javascript:","")
  eval(jscmd)
  }
}