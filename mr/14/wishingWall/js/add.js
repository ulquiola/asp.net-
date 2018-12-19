//关闭贴字条窗口
function close_window(){
	scripForm.wishMan.value="";
	scripForm.wellWisher.value="匿名";
	document.getElementById("preview").setAttribute("class","Style0");
	document.getElementById("preview").setAttribute("className","Style0");
	scripForm.color.value=0;
	scripForm.face.value=0;
	document.getElementById("pFace").src = "images/face/face_0.gif";	
	scripForm.content.value="";
	document.getElementById("pWishMan").innerHTML="";
	document.getElementById("pWellWisher").innerHTML="匿名";	
	document.getElementById("pContext").innerHTML="";
	scripForm.used.value=0;
	scripForm.remain.value=200;		//初始化最大字节数
	document.getElementById("preview").style.display = "none";	
	document.getElementById("scrip_add").style.display="none";
	document.getElementById("notClickDiv").style.display ="none";
	document.getElementById("yanzheng").value = "";
	document.getElementById("chknm").value = "";
	
}
//打开贴字条窗口
function loadScripAdd_window(){
	/**********显示其他内容不允许点击的图层*****************/
	document.getElementById("notClickDiv").style.display = "block";
	var w =document.body.clientWidth;
	var h= document.body.clientHeight;
	document.getElementById("notClickDiv").style.width = w;
	document.getElementById("notClickDiv").style.height = h;	
	rsCount=document.getElementById("hRsCount").value;
	document.getElementById("notClickDiv").style.zIndex=rsCount+101;
	/**********显示添加字条的图层*****************/
	var width=screen.width;
	var height=screen.height;
	document.getElementById("scrip_add").style.display = "block";
	document.getElementById("scrip_add").style.top=(height-310-240)/2;
	document.getElementById("scrip_add").style.left=(width-500)/2-120;
	document.getElementById("scrip_add").style.zIndex=rsCount+102;
	/**********显示预览字条的图层*****************/
	document.getElementById("preview").style.display = "block";
	document.getElementById("preview").setAttribute("class","Style0");
	document.getElementById("preview").style.top=(height-310-240)/2;
	document.getElementById("preview").style.left=(width-240)/2+290;
	document.getElementById("preview").style.zIndex=rsCount+103;	
	showval();
}
function InputInfo(OriInput, GoalArea){
	document.getElementById(GoalArea).innerHTML = OriInput.value;
}

/*******************限制字条字节数***********************************************/
 var LastCount =0;
 function CountStrByte(info,Total,Used,Remain){ //字节统计
 var ByteCount = 0;
 var StrValue  = info.value;
 var StrLength = info.value.length;
 var MaxValue  = Total.value;

 if(LastCount != StrLength) { // 在此判断，减少循环次数
	for (i=0;i<StrLength;i++){
		ByteCount  = (StrValue.charCodeAt(i)<=256) ? ByteCount + 1 : ByteCount + 2;
      if (ByteCount>MaxValue) {
      	info.value = StrValue.substring(0,i);
			alert("留言内容最多不能超过 " +MaxValue+ " 个字节！\n注意：一个汉字为两字节。");
         ByteCount = MaxValue;
         break;
      }
	}
   Used.value = ByteCount;
   Remain.value = MaxValue - ByteCount;
   LastCount = StrLength;
   document.getElementById("pContext").innerHTML = StrValue;
 }
}
/******************************************************************************/
function ColorChoose(n){
	//修改字条背景
	var ClassName = "Style"+n;
	document.getElementById("preview").setAttribute("class",ClassName);
	document.getElementById("preview").setAttribute("className",ClassName);
	scripForm.color.value = n;
}
function faceChoose(n){
	//修改心情图标
	var Url = "images/face/face_"+n+".gif";
	document.getElementById("pFace").src = Url;	
	scripForm.face.value = n;
}

function getTime(){
	//获得当前时间
	var ThisTime = new Date();
	return ThisTime.getFullYear()+"-"+(ThisTime.getMonth()+1)+"-"+ThisTime.getDate(); 
}
//***************************************************保存字条********************************************************/
function scripSubmit(){
	if(scripForm.wishMan.value==""){
		alert("请输入祝福对象！");scripForm.wishMan.focus();return false;
	}
	if(scripForm.wellWisher.value==""){
		scripForm.wellWisher.value="匿名";
	}
	if(scripForm.content.value==""){
		alert("请输入字条内容！");scripForm.content.focus();return false;
	}
	
	if(scripForm.yanzheng.value==""){
		alert("验证码不能为空！");scripForm.yanzheng.focus();return false;
	}
	if(scripForm.yanzheng.value != document.getElementById("chknm").value){
		alert("验证码输入错误！");
		return false;
	}

	var param="wishMan="+scripForm.wishMan.value+"&wellWisher="+scripForm.wellWisher.value+"&color="+scripForm.color.value+"&face="+scripForm.face.value+"&content="+scripForm.content.value; 

	var loader=new net.AjaxRequest("wishadd_deal.asp?"+param,deal_s,onerror,"GET",param);
}


///////////////////////////////
function showval(){
		num = '';
		for(i=0;i<4;i++){
			tmp =  Math.ceil((Math.random() * 9));
				num += tmp;
		}
		document.getElementById('sss').src='js/yz.asp?num='+num;
		document.getElementById('chknm').value = num;
}

/////////////////////////////////




function deal_s(){
	/***获取字条编号*/
	var h=this.req.responseText;
	alert(h);
	h=h.replace(/\s/g,"");	//去除字符串中的Unicode空白符
	var id=h.substr(h.indexOf("[")+1,h.indexOf("]")-h.indexOf("[")-1);
	
	/****************/
	if(h!="很报谦，您的字条发送失败！"){
		createNewScrip(id);		//显示该字条
		Show(id,'shadeDiv');	//设置新添加的字条突出显示
	}
	close_window();	//关闭添加字条窗口
}
