    <script language="javascript">
			    function chkinput(form){
				 if(form.recuser.value==""){
				   alert('�������ջ���������');
				   form.recuser.focus();
				   return(false);
				 }
				 
				 if(form.address.value==""){
				   alert('�������ջ�����ϵ��ַ��');
				   form.address.focus();
				   return(false);
				 }
				  
				 if(form.yb.value==""){
				   alert('�������ջ����ʱ࣡');
				   form.yb.focus();
				   return(false);
				 }
				 
				 if(isNaN(form.yb.value)){
				   alert('�ʱ�ֻ����������ɣ�');
				   form.yb.focus();
				   return(false);
				 }
				 
				 if(form.yb.value.length!=6){
				   alert('�ʱ����6λ��ɣ�');
				   form.yb.focus();
				   return(false);
				 }
				 
				 
				 if(form.qq.value==""){
			
			   alert("����д����QQ���룡");
			   form.qq.select();
			   return(false);
			
			   }
		   
		      if(isNaN(form.qq.value)==true){
			
			   alert("QQ��ֻ����������ɣ�");
			   form.qq.select();
			   return(false);
			}
				 
				 
				 
				if(form.email.value=="")
	          {
	             alert("������E-mail��ַ!");
	             form.email.select();
	             return(false);
	            }
	        var i=form.email.value.indexOf("@");
	        var j=form.email.value.indexOf(".");
	       if((i<0)||(i-j>0)||(j<0))
	        {
              alert("��������ȷ��E-mail��ַ!");
	          form.email.select();
	          return(false);
	        } 
				
				
				
				 if(form.mtel.value==""){
				   alert('�������ƶ��绰���룡');
				   form.mtel.focus();
				   return(false);
				 }
				 
			  	 if(form.gtel.value==""){
				   alert('������̶��绰���룡');
				   form.gtel.focus();
				   return(false);
				 }
	 
				 
			  return(true);
				
			}
			  
			  </script>
			  	  