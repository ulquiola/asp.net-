// JavaScript Document
function $(id){
	return document.getElementById(id);
}
window.onload = function(){
	$('regname').focus();
	var cname1,cname2,cpwd1,cpwd2;
	//���ü��ť
	function chkreg(){
		if((cname1 == 'yes') && (cname2 == 'yes') && (cpwd1 == 'yes') && (cpwd2 == 'yes')){
			$('regbtn').disabled = false;
		}else{
			$('regbtn').disabled = true;
		}
	}
	//��֤�û���
	$('regname').onkeyup = function (){
		name = $('regname').value;
		cname2 = '';
		if(name != '' && name.match(/^[a-zA-Z_]*/) == ''){
			$('namediv').innerHTML = '<font color=red>��������ĸ���»��߿�ͷ</font>';
			cname1 = '';
		}else if(name != '' && name.length < 2){
			$('namediv').innerHTML = '<font color=red>ע�����Ʊ�����ڵ���2λ</font>';
			cname1 = '';
		}else if(name == ''){
			$('namediv').innerHTML = '';
		}else{
			$('namediv').innerHTML = '<font color=green>ע�����Ʒ��ϱ�׼</font>';
			cname1 = 'yes';
		}
		chkreg();
	}
	//��֤�Ƿ���ڸ��û�
	$('regname').onblur = function(){
		name = $('regname').value;
		if(cname1 == 'yes' && name != ''){
			xmlhttp.open('get','chkname.asp?name='+name,true);
			xmlhttp.onreadystatechange = function(){
				if(xmlhttp.readyState == 4){
					if(xmlhttp.status == 200){
						var msg = xmlhttp.responseText;
						$('namediv').innerHTML = msg;
						if(msg == '2'){
							$('namediv').innerHTML="<font color=green>��ϲ�������û�������ʹ��!</font>";
							cname2 = 'yes';
						}else if(msg == '1'){
							$('namediv').innerHTML="<font color=red>�û�����ռ�ã�</font>";
							cname2 = '';
						}else if(msg == '3'){
							$('namediv').innerHTML = '<font color=red>�û�������Ϊ��!</font>';
							cname2= '';
						}else{
							$('namediv').innerHTML="<font color=red>"+msg+"</font>";
							cname2 = '';
						}
					}
				}
				chkreg();
			}
			xmlhttp.send(null);
		}
	}
	//��֤����
	$('regpwd1').onkeyup = function(){
			pwd = $('regpwd1').value;
			pwd2 = $('regpwd2').value;
		if(pwd == ''){
			$('pwddiv1').innerHTML = '';
			cpwd1 = '';
		}else if(pwd.length < 6 && pwd != ''){
			$('pwddiv1').innerHTML = '<font color=red>���볤��������Ҫ6λ</font>';
			cpwd1 = '';
		}else if(pwd.length >= 6 && pwd.length < 12){
			$('pwddiv1').innerHTML = '<font color=green>�������Ҫ������ǿ�ȣ���</font>';
			cpwd1 = 'yes';
		}else if((pwd.match(/^[0-9]*$/)!=null) || (pwd.match(/^[a-zA-Z]*$/) != null )){
			$('pwddiv1').innerHTML = '<font color=green>�������Ҫ������ǿ�ȣ���</font>';
			cpwd1 = 'yes';
		}else{
			$('pwddiv1').innerHTML = '<font color=green>�������Ҫ������ǿ�ȣ���</font>';
			cpwd1 = 'yes';
		}
		if(pwd2 != '' && pwd != pwd2){
			$('pwddiv2').innerHTML = '<font color=red>�������벻һ��!</font>';
			cpwd2 = '';
		}else if(pwd2 != '' && pwd == pwd2){
			$('pwddiv2').innerHTML = '<font color=green>����������ȷ</font>';
			cpwd2 = 'yes';
		}
		chkreg();
	}
	//��֤ȷ������
	$('regpwd2').onkeyup = function(){
		pwd1 = $('regpwd1').value;
		pwd2 = $('regpwd2').value;
		if(pwd1 == '' || pwd2 == ''){
			$('pwddiv2').innerHTML = '';
			cpwd2 = '';
		}else if(pwd1 != pwd2){
			$('pwddiv2').innerHTML = '<font color=red>�������벻һ��!</font>';
			cpwd2 = '';
		}else{
			$('pwddiv2').innerHTML = '<font color=green>����������ȷ</font>';
			cpwd2 = 'yes';
			chkreg();
		}
	}
	//��ʾ/������ϸ��Ϣ
	$('morebtn').onclick = function(){
		
		if($('morediv').style.display == ''){
			$('morediv').style.display = 'none';
		}else{
			$('morediv').style.display = '';
		}
	}
	//��¼��ť
	$('logbtn').onclick = function(){
		window.open('login.asp','_parent','',false);
	}
	//��ʽע��
	$('regbtn').onclick = function(){
		$('imgdiv').style.visibility = 'visible';
		url = 'register_chk.asp?name='+$('regname').value+'&pwd='+$('regpwd1').value;
		url += '&question=' +$('question').value+'&answer='+$('answer').value;
		url += '&realname=' +$('realname').value+'&email='+$('email').value;
		url += '&telephone='+$('telephone').value+'&qq='+$('qq').value;
		xmlhttp.open('get',url,true);
		xmlhttp.onreadystatechange = function(){
			if(xmlhttp.readyState == 4){
				if(xmlhttp.status == 200){
					msg = xmlhttp.responseText;
					//alert(msg);
					$('pwddiv2').innerHTML=msg;
					if(msg == '1'){
						alert('ע��ɹ�,');
						location='main.asp';
					}else{
						alert(msg);
					}
					$('imgdiv').style.visibility = 'hidden';
				}
			}
		}
		xmlhttp.send(null);
	}
}