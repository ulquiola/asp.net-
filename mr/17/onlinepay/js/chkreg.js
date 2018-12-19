function checkregtel(regtel){
	var str=regtel;
	var Expression=/^13(\d{9})$|^15(\d{9})$/; 
	var objExp=new RegExp(Expression);
	if(objExp.test(str)==true){
		return true;
	}else{
		return false;
	}
}





function checkregtels(regtels){
	var str=regtels;
	var Expression=/^(\d{3}-)(\d{8})$|^(\d{4}-)(\d{7})$|^(\d{4}-)(\d{8})$/; 
	var objExp=new RegExp(Expression);
	if(objExp.test(str)==true){
		return true;
	}else{
		return false;
	}
}


function chkreginfo(form,mark,edit){
 

  if(mark==0 || mark=="all"){
   	 if(form.recuser.value==""){
	   chknew_recuser.innerHTML="�������ջ��ˣ�";
	   form.recuser.style.backgroundColor="#FF0000";
	   return false;
   	 }else{
   	   chknew_recuser.innerHTML="";
	   form.recuser.style.backgroundColor="#FFFFFF";
   	 }
   }
   

  if(mark==1 || mark=="all"){
   	 if(form.address.value==""){
	   chknew_address.innerHTML="��������ϵ��ַ��";
	   form.address.style.backgroundColor="#FF0000";
	   return false;
   	 }else{
   	   chknew_address.innerHTML="";
	   form.address.style.backgroundColor="#FFFFFF";
   	 }
   }


 if(mark==2 || mark=="all"){
   	 if(form.yb.value==""){
	   chknew_yb.innerHTML="�������ʱ࣡";
	   form.yb.style.backgroundColor="#FF0000";
	   return false;
   	 }else if(isNaN(form.yb.value)){
   	   chknew_yb.innerHTML="�ʱ���������ɣ�";
	   form.yb.style.backgroundColor="#FF0000";
	   return false;
	 }else if(form.yb.value.length!=6){
   	   chknew_yb.innerHTML="�ʱ���6λ������ɣ�";
	   form.yb.style.backgroundColor="#FF0000";
	   return false;
   	 }else{	
   	   chknew_yb.innerHTML="";
	   form.yb.style.backgroundColor="#FFFFFF";
   	 }
   }
   
   
   if(mark==3 || mark=="all"){
   	 if(form.qq.value==""){
	   chknew_qq.innerHTML="������QQ���룡";
	   form.qq.style.backgroundColor="#FF0000";
	   return false;
   	 }else if(isNaN(form.qq.value)){
   	   chknew_qq.innerHTML="QQ����������ɣ�";
	   form.qq.style.backgroundColor="#FF0000";
	   return false;
   	 }else{	
   	   chknew_qq.innerHTML="";
	   form.qq.style.backgroundColor="#FFFFFF";
   	 }
   }



  if(mark==4 || mark=="all"){
   	 if(form.email.value==""){
	   chknew_email.innerHTML="������E-mail��ַ��";
	   form.email.style.backgroundColor="#FF0000";
	   return false; 
   	 }else{	
   	   var i=form.email.value.indexOf("@");
	   var j=form.email.value.indexOf(".");
	   if((i<0)||(i-j>0)||(j<0)){
         chknew_email.innerHTML="E-mail��ַ��ʽ����";
	     form.email.style.backgroundColor="#FF0000";
	     return false;                              
	   }else{
	   	  chknew_email.innerHTML="";
	      form.email.style.backgroundColor="#FFFFFF";
	   }
	   
   	 }
   }
   

  if(mark==5 || mark=="all"){
		if(form.mtel.value==""){
	   		chknew_mtel.innerHTML="������绰���룡";
	   		form.mtel.style.backgroundColor="#FF0000";
	   		return false;
   	 	}else if(!checkregtel(form.mtel.value)){
	   		chknew_mtel.innerHTML="�绰����ĸ�ʽ����ȷ��";
	   		form.mtel.style.backgroundColor="#FF0000";
	   		return false;
   	 	}else if(isNaN(form.mtel.value)){
   	   		chknew_mtel.innerHTML="�绰����������ɣ�";
	   		form.mtel.style.backgroundColor="#FF0000";
	   		return false;
   	 	}else{	
   	   		chknew_mtel.innerHTML="";
	   		form.mtel.style.backgroundColor="#FFFFFF";
   	 	}
   	}


   if(mark==6 || mark=="all"){
		if(form.gtel.value==""){
	   		chknew_gtel.innerHTML="������绰���룡";
	   		form.gtel.style.backgroundColor="#FF0000";
	   		return false;
   	 	}else if(!checkregtels(form.gtel.value)){
	   		chknew_gtel.innerHTML="�绰����ĸ�ʽ����ȷ��";
	   		form.gtel.style.backgroundColor="#FF0000";
	   		return false;
   	 	}else{	
   	   		chknew_gtel.innerHTML="";
	   		form.gtel.style.backgroundColor="#FFFFFF";
   	 	}
   	}


   }


