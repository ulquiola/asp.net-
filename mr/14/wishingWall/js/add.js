//�ر�����������
function close_window(){
	scripForm.wishMan.value="";
	scripForm.wellWisher.value="����";
	document.getElementById("preview").setAttribute("class","Style0");
	document.getElementById("preview").setAttribute("className","Style0");
	scripForm.color.value=0;
	scripForm.face.value=0;
	document.getElementById("pFace").src = "images/face/face_0.gif";	
	scripForm.content.value="";
	document.getElementById("pWishMan").innerHTML="";
	document.getElementById("pWellWisher").innerHTML="����";	
	document.getElementById("pContext").innerHTML="";
	scripForm.used.value=0;
	scripForm.remain.value=200;		//��ʼ������ֽ���
	document.getElementById("preview").style.display = "none";	
	document.getElementById("scrip_add").style.display="none";
	document.getElementById("notClickDiv").style.display ="none";
	document.getElementById("yanzheng").value = "";
	document.getElementById("chknm").value = "";
	
}
//������������
function loadScripAdd_window(){
	/**********��ʾ�������ݲ���������ͼ��*****************/
	document.getElementById("notClickDiv").style.display = "block";
	var w =document.body.clientWidth;
	var h= document.body.clientHeight;
	document.getElementById("notClickDiv").style.width = w;
	document.getElementById("notClickDiv").style.height = h;	
	rsCount=document.getElementById("hRsCount").value;
	document.getElementById("notClickDiv").style.zIndex=rsCount+101;
	/**********��ʾ���������ͼ��*****************/
	var width=screen.width;
	var height=screen.height;
	document.getElementById("scrip_add").style.display = "block";
	document.getElementById("scrip_add").style.top=(height-310-240)/2;
	document.getElementById("scrip_add").style.left=(width-500)/2-120;
	document.getElementById("scrip_add").style.zIndex=rsCount+102;
	/**********��ʾԤ��������ͼ��*****************/
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

/*******************���������ֽ���***********************************************/
 var LastCount =0;
 function CountStrByte(info,Total,Used,Remain){ //�ֽ�ͳ��
 var ByteCount = 0;
 var StrValue  = info.value;
 var StrLength = info.value.length;
 var MaxValue  = Total.value;

 if(LastCount != StrLength) { // �ڴ��жϣ�����ѭ������
	for (i=0;i<StrLength;i++){
		ByteCount  = (StrValue.charCodeAt(i)<=256) ? ByteCount + 1 : ByteCount + 2;
      if (ByteCount>MaxValue) {
      	info.value = StrValue.substring(0,i);
			alert("����������಻�ܳ��� " +MaxValue+ " ���ֽڣ�\nע�⣺һ������Ϊ���ֽڡ�");
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
	//�޸���������
	var ClassName = "Style"+n;
	document.getElementById("preview").setAttribute("class",ClassName);
	document.getElementById("preview").setAttribute("className",ClassName);
	scripForm.color.value = n;
}
function faceChoose(n){
	//�޸�����ͼ��
	var Url = "images/face/face_"+n+".gif";
	document.getElementById("pFace").src = Url;	
	scripForm.face.value = n;
}

function getTime(){
	//��õ�ǰʱ��
	var ThisTime = new Date();
	return ThisTime.getFullYear()+"-"+(ThisTime.getMonth()+1)+"-"+ThisTime.getDate(); 
}
//***************************************************��������********************************************************/
function scripSubmit(){
	if(scripForm.wishMan.value==""){
		alert("������ף������");scripForm.wishMan.focus();return false;
	}
	if(scripForm.wellWisher.value==""){
		scripForm.wellWisher.value="����";
	}
	if(scripForm.content.value==""){
		alert("�������������ݣ�");scripForm.content.focus();return false;
	}
	
	if(scripForm.yanzheng.value==""){
		alert("��֤�벻��Ϊ�գ�");scripForm.yanzheng.focus();return false;
	}
	if(scripForm.yanzheng.value != document.getElementById("chknm").value){
		alert("��֤���������");
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
	/***��ȡ�������*/
	var h=this.req.responseText;
	alert(h);
	h=h.replace(/\s/g,"");	//ȥ���ַ����е�Unicode�հ׷�
	var id=h.substr(h.indexOf("[")+1,h.indexOf("]")-h.indexOf("[")-1);
	
	/****************/
	if(h!="�ܱ�ǫ��������������ʧ�ܣ�"){
		createNewScrip(id);		//��ʾ������
		Show(id,'shadeDiv');	//��������ӵ�����ͻ����ʾ
	}
	close_window();	//�ر������������
}
