function open_close(x){
  
   if(x.style.display==''){
	  x.style.display='none'; 
    }else if(x.style.display=='none'){
	  x.style.display=''; 

    }
   
}


function open_close1(x,y,z){
 
   if(x.style.display==''){
	  x.style.display='none'; 
	  y.src='images/tplus.gif';
	  z.src='images/1.gif';
    }else if(x.style.display=='none'){
	  x.style.display=''; 
	  y.src='images/tplus1.gif';	
	  z.src='images/2.gif';
    }
	
}

function open_close2(x,y,z){
 
   if(x.style.display==''){
	  x.style.display='none'; 
	  y.src='images/uaplus.gif';
	  z.src='images/1.gif';
    }else if(x.style.display=='none'){
	  x.style.display=''; 
	  y.src='images/uaplus1.gif';	
	  z.src='images/2.gif';
    }
	
}

function open_close3(objImg,objTr){
	if(objTr.style.display == ""){
		objTr.style.display = "none";
		objImg.src = "images/jia.gif";
		objImg.alt = "չ��";		
	}else if(objTr.style.display == "none"){
		objTr.style.display = "";
		objImg.src = "images/jian.gif";
		objImg.alt = "�۵�";		
	}
}


function submitu(form){
   if(form.usernc.value==""){
     alert("�����������ǳ�");
	 form.usernc.select();
	 return(false);
	}
	if(form.userpwd.value==""){
     alert("�������¼���룡");
	 form.userpwd.select();
	 return(false);
    }
	if(form.xym.value==""){
     alert("��������֤�룡");
	 form.xym.select();
	 return(false);
    }
	if(form.xym.value!=form.xym1.value){
     alert("��֤���������");
	 form.xym.select();
	 return(false);
    }
	
   return(true);	 
 }
 
 
 	  function getStart(textBox){
    
        if(typeof(textBox.selectionStart) == "number"){
            start = textBox.selectionStart;
        }
       
        else if(document.selection){
            var range = document.selection.createRange();
            if(range.parentElement().id == textBox.id){
            
                var range_all = document.body.createTextRange();
                range_all.moveToElementText(textBox);
              
                for (start=0; range_all.compareEndPoints("StartToStart", range) < 0; start++)
                    range_all.moveStart('character', 1);
              
                for (var i = 0; i <= start; i ++){
                    if (textBox.value.charAt(i) == '\n')
                        start++;
                }
              
              }
           }
     
	  return(start);
	 
	}
	
	
	
		function getEnd(textBox){
    
        if(typeof(textBox.selectionStart) == "number"){
         
            end = textBox.selectionEnd;
        }
       
        else if(document.selection){
            var range = document.selection.createRange();
            if(range.parentElement().id == textBox.id){
            
                 var range_all = document.body.createTextRange();
                 range_all.moveToElementText(textBox);
           
                 for (end = 0; range_all.compareEndPoints('StartToEnd', range) < 0; end ++)
                     range_all.moveStart('character', 1);
             
                     for (var i = 0; i <= end; i ++){
                         if (textBox.value.charAt(i) == '\n')
                             end ++;
                     }
                }
            }
      
     
	  return(end);
	}
	


      
	 function connectText(text,start,end,markstart,markend){
	 
	     if(end==0 || start=="undefined" || end=="undefined"){
		    alert("��ѡ��Ҫ�������õ����ݣ�");
			return;	 
		 }
		 t=text.value; 
		 t1=t.substr(0,start);
		 t2=markstart+t.substring(start,end)+markend;
		 t3=t.substring(end,t.length);
		 text.value=(t1+t2+t3);
	 
	 } 
	 
	 function addImage(text,address){
	 	if(address=="" || address=="http://"){
	 	  alert("������ͼƬ��ַ��");
	 	}else
	 	  text.value=text.value+'[img src='+address+'/img]';
	 }
	
	
	 function chkinputbbs(form){
	    
		  if(form.title.value==""){
			  
			alert("�������������⣡");
			form.title.focus();
			return(false);
		  }
		  
		  if(form.content.value==""){
			alert("�������������ݣ�");  
			form.content.focus();
			return(false);
		  }
		  
		if(form.xym.value==""){
             alert("��������֤�룡");
	         form.xym.select();
	         return(false);
        }
	    if(form.xym.value!=form.xym1.value){
           alert("��֤���������");
	       form.xym.select();
	       return(false);
        }
		 
		 return(true);
		 
     }
	 
      function chkinputeditbbs(form){
     	if(form.title.value==""){
			  
			alert("�������������⣡");
			form.title.focus();
			return(false);
		  }
		  
		  if(form.content.value==""){
			alert("�������������ݣ�");  
			form.content.focus();
			return(false);
		  }
     	return true;
     }
     
     function chkinputeditreply(form){
     	if(form.title.value==""){
			  
			alert("������ظ����⣡");
			form.title.focus();
			return(false);
		  }
		  
		  if(form.content.value==""){
			alert("������ظ����ݣ�");  
			form.content.focus();
			return(false);
		  }
     	return true;
     }
     
	 	 function chkinputreply(form){
	    
		  if(form.title.value==""){
			  
			alert("������ظ����⣡");
			form.title.focus();
			return(false);
		  }
		  
		  if(form.content.value==""){
			alert("������ظ����ݣ�");  
			form.content.focus();
			return(false);
		  }
		  if(form.xym.value==""){
             alert("��������֤�룡");
	         form.xym.select();
	         return(false);
        }
	    if(form.xym.value!=form.xym1.value){
           alert("��֤���������");
	       form.xym.select();
	       return(false);
        }
		 
		 return(true);
		 
     }

     function chkinputjb(form){
	    
		  if(form.title.value==""){
			  
			alert("������ٱ�ԭ��");
			form.title.focus();
			return(false);
		  }
		  
		  if(form.content.value==""){
			alert("������������ݣ�");  
			form.content.focus();
			return(false);
		  }
		  if(form.xym.value==""){
             alert("��������֤�룡");
	         form.xym.select();
	         return(false);
        }
	    if(form.xym.value!=form.xym1.value){
           alert("��֤���������");
	       form.xym.select();
	       return(false);
        }
		 
		 return(true);
		 
     }


   
   function changtextarea(x,y){
	 if(y=="+"){  
	   x.rows=x.rows+10;
	  }else if(y=="-"){
	    x.rows=x.rows-10; 
	  }
   }
   
   function showreplyform(){
	    
	  if(rform1.style.display=='none'){
		  
	     rform1.style.display='';
		 rform2.style.display='';
	  }else if(rform1.style.display==''){
		 rform1.style.display='none';
		 rform2.style.display='none';
		  
	  }   
	   
   }
   
    
   function additem(form){
   	 if(form.sex.value=="��"){
   	 	for(var i=6;i<=11;i++){
   	 		form.photo.options[i-6]=new Option(i+".gif");	
   	 	}
   	 	form.img.src="images/face/6.gif"
   	 }else if(form.sex.value=="Ů"){
   	 	for(var i=0;i<=5;i++){
   	 		form.photo.options[i]=new Option(i+".gif");
   	 	}
   	 	form.img.src="images/face/0.gif"
   	 } 
   	
   	
   } 
   
   
   function selectitem(form){
   	 if(form.sex.value=="��"){
   	 	form.img.src='images/face/'+(form.photo.selectedIndex+6)+'.gif'
   	 }else{
   	 	form.img.src='images/face/'+form.photo.selectedIndex+'.gif'
   	 }
   	
   }
   
  function chkinputfindbbs(form){
  	if(form.keyword.value==""){
  		alert("�������ѯ�ؼ���");
  		form.keyword.focus();
  		return false;
  	}
  	return true;
  }
  
  
  
//�ж��û��������Ƿ�Ϸ�
function check(){
	if (myform.UserName.value==""){
		alert("�������û�����");myform.UserName.focus();return;
	}
	if (myform.TrueName.value==""){
		alert("��������ʵ������");myform.TrueName.focus();return;
	}
	if (myform.PWD1.value==""){
		alert("���������룡");myform.PWD1.focus();return;
	}
	if (myform.PWD1.value.length<6){
		alert("��������Ϊ6λ�����������룡");myform.PWD1.focus();return;
	}		
	if (myform.PWD2.value==""){
		alert("��ȷ�����룡");myform.PWD2.focus();return;
	}
	if (myform.PWD1.value!=myform.PWD2.value){
		alert("��������������벻һ�£����������룡");myform.PWD1.focus();return;
	}
	if(myform.birthday.value==""){
		alert("��������������");myform.birthday.focus();return;
	}		
	if(CheckDate(myform.birthday.value)){
		alert("�������׼���ڣ��磺1990/03/13��1990-03-13��");myform.birthday.focus();return;
	}
	if (myform.Email.value==""){
		alert("������Email��ַ��");myform.Email.focus();return;
	}
	var i=myform.Email.value.indexOf("@");
	var j=myform.Email.value.indexOf(".");
	if((i<0)||(i-j>0)||(j<0)){
		alert("�������Email��ַ����ȷ�����������룡");myform.Email.value="";myform.Email.focus();return;
	}
	myform.submit();		
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
   return false;}
