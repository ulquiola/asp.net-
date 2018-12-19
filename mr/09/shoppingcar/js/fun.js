    <script language="javascript">
			    function chkinput(form){
				 if(form.recuser.value==""){
				   alert('请输入收货人姓名！');
				   form.recuser.focus();
				   return(false);
				 }
				 
				 if(form.address.value==""){
				   alert('请输入收货人联系地址！');
				   form.address.focus();
				   return(false);
				 }
				  
				 if(form.yb.value==""){
				   alert('请输入收货人邮编！');
				   form.yb.focus();
				   return(false);
				 }
				 
				 if(isNaN(form.yb.value)){
				   alert('邮编只能由数字组成！');
				   form.yb.focus();
				   return(false);
				 }
				 
				 if(form.yb.value.length!=6){
				   alert('邮编号由6位组成！');
				   form.yb.focus();
				   return(false);
				 }
				 
				 
				 if(form.qq.value==""){
			
			   alert("请填写您的QQ号码！");
			   form.qq.select();
			   return(false);
			
			   }
		   
		      if(isNaN(form.qq.value)==true){
			
			   alert("QQ号只能由数字组成！");
			   form.qq.select();
			   return(false);
			}
				 
				 
				 
				if(form.email.value=="")
	          {
	             alert("请输入E-mail地址!");
	             form.email.select();
	             return(false);
	            }
	        var i=form.email.value.indexOf("@");
	        var j=form.email.value.indexOf(".");
	       if((i<0)||(i-j>0)||(j<0))
	        {
              alert("请输入正确的E-mail地址!");
	          form.email.select();
	          return(false);
	        } 
				
				
				
				 if(form.mtel.value==""){
				   alert('请输入移动电话号码！');
				   form.mtel.focus();
				   return(false);
				 }
				 
			  	 if(form.gtel.value==""){
				   alert('请输入固定电话号码！');
				   form.gtel.focus();
				   return(false);
				 }
	 
				 
			  return(true);
				
			}
			  
			  </script>
			  	  