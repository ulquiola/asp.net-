var Layer='';
var a;
document.onmouseup=moveEnd;
document.onmousemove=moveStart;
var b;
var c;
function Move(Object,event){
	Layer=Object.id;
	if(document.all){
		document.getElementById(Layer).setCapture();
		b=event.x-document.getElementById(Layer).style.pixelLeft;
		c=event.y-document.getElementById(Layer).style.pixelTop;
	}else if(window.captureEvents){
		window.captureEvents(Event.MOUSEMOVE|Event.MOUSEUP);
		b=event.layerX;
		c=event.layerY;
	}
	/**实现鼠标单击字条时，字条置上**/
	if(Layer!="scrip"+a){
		document.getElementById(Layer).style.zIndex=iLayerMaxNum;
		iLayerMaxNum=iLayerMaxNum+1;
	}
	/********************************/
}
function moveStart(d){
	if(Layer!=''){
		if(document.all){
			document.getElementById(Layer).style.left=event.x-b;
			document.getElementById(Layer).style.top=event.y-c;
		}else if(window.captureEvents){
			document.getElementById(Layer).style.left=(d.clientX-b)+"px";
			document.getElementById(Layer).style.top=(d.clientY-c)+"px";
		}
	}
}
function moveEnd(d){
	if(Layer!=''){
		if(document.all){
			document.getElementById(Layer).releaseCapture();
			Layer='';
		}else if(window.captureEvents){
			window.captureEvents(Event.MOUSEMOVE|Event.MOUSEUP);
			Layer='';
		}
	}
}
function myClose(n){
	var e='scrip'+n;											
	document.getElementById(e).style.display='none';
}
function Show(n,divName){
	var e='scrip'+n;
	document.getElementById(e).style.zIndex =iLayerMaxNum+1;
	document.getElementById(divName).style.display = "block";
	document.getElementById(divName).style.zIndex = iLayerMaxNum;
	document.getElementById(divName).style.width = document.body.clientWidth;
	document.getElementById(divName).style.height = document.body.clientHeight;	
}	
function Hide(){
	document.getElementById("shadeDiv").style.display = "none";
	iLayerMaxNum=iLayerMaxNum+2;
}
/***************************************添加人气*****************************************************************/
function holdout(id){
	if(id>0){
		var loader=new net.AjaxRequest("zhichi_deal.asp?scripId="+id+"&nocache="+new Date().getTime(),deal_addHoldout,onerror,"GET");
	}
}
function deal_addHoldout(){
	h=this.req.responseText;
	var hitsId="hitsValue"+h.substr(0,h.indexOf("#"));
	hitsId=hitsId.replace(/\s/g,"");	//去除字符串中的Unicode空白符
	var hitsNum=h.substr(h.indexOf("#")+1,h.length-h.indexOf("#")-1);	
	document.getElementById(hitsId).innerHTML=hitsNum;
}
/*******************************************************************************************************************/