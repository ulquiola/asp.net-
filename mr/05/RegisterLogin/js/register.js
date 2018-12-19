// JavaScript Document
function $(id){
	return document.getElementById(id);
}
window.onload = function(){
	$('regname').focus();
	var cname1,cname2,cpwd1,cpwd2;
	//设置激活按钮
	function chkreg(){
		if((cname1 == 'yes') && (cname2 == 'yes') && (cpwd1 == 'yes') && (cpwd2 == 'yes')){
			$('regbtn').disabled = false;
		}else{
			$('regbtn').disabled = true;
		}
	}
	//验证用户名
	$('regname').onkeyup = function (){
		name = $('regname').value;
		cname2 = '';
		if(name != '' && name.match(/^[a-zA-Z_]*/) == ''){
			$('namediv').innerHTML = '<font color=red>必须以字母或下划线开头</font>';
			cname1 = '';
		}else if(name != '' && name.length < 2){
			$('namediv').innerHTML = '<font color=red>注册名称必须大于等于2位</font>';
			cname1 = '';
		}else if(name == ''){
			$('namediv').innerHTML = '';
		}else{
			$('namediv').innerHTML = '<font color=green>注册名称符合标准</font>';
			cname1 = 'yes';
		}
		chkreg();
	}
	//验证是否存在该用户
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
							$('namediv').innerHTML="<font color=green>恭喜您，该用户名可以使用!</font>";
							cname2 = 'yes';
						}else if(msg == '1'){
							$('namediv').innerHTML="<font color=red>用户名被占用！</font>";
							cname2 = '';
						}else if(msg == '3'){
							$('namediv').innerHTML = '<font color=red>用户名不能为空!</font>';
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
	//验证密码
	$('regpwd1').onkeyup = function(){
			pwd = $('regpwd1').value;
			pwd2 = $('regpwd2').value;
		if(pwd == ''){
			$('pwddiv1').innerHTML = '';
			cpwd1 = '';
		}else if(pwd.length < 6 && pwd != ''){
			$('pwddiv1').innerHTML = '<font color=red>密码长度最少需要6位</font>';
			cpwd1 = '';
		}else if(pwd.length >= 6 && pwd.length < 12){
			$('pwddiv1').innerHTML = '<font color=green>密码符合要求。密码强度：弱</font>';
			cpwd1 = 'yes';
		}else if((pwd.match(/^[0-9]*$/)!=null) || (pwd.match(/^[a-zA-Z]*$/) != null )){
			$('pwddiv1').innerHTML = '<font color=green>密码符合要求。密码强度：中</font>';
			cpwd1 = 'yes';
		}else{
			$('pwddiv1').innerHTML = '<font color=green>密码符合要求。密码强度：高</font>';
			cpwd1 = 'yes';
		}
		if(pwd2 != '' && pwd != pwd2){
			$('pwddiv2').innerHTML = '<font color=red>两次密码不一致!</font>';
			cpwd2 = '';
		}else if(pwd2 != '' && pwd == pwd2){
			$('pwddiv2').innerHTML = '<font color=green>密码输入正确</font>';
			cpwd2 = 'yes';
		}
		chkreg();
	}
	//验证确认密码
	$('regpwd2').onkeyup = function(){
		pwd1 = $('regpwd1').value;
		pwd2 = $('regpwd2').value;
		if(pwd1 == '' || pwd2 == ''){
			$('pwddiv2').innerHTML = '';
			cpwd2 = '';
		}else if(pwd1 != pwd2){
			$('pwddiv2').innerHTML = '<font color=red>两次密码不一致!</font>';
			cpwd2 = '';
		}else{
			$('pwddiv2').innerHTML = '<font color=green>密码输入正确</font>';
			cpwd2 = 'yes';
			chkreg();
		}
	}
	//显示/隐藏详细信息
	$('morebtn').onclick = function(){
		
		if($('morediv').style.display == ''){
			$('morediv').style.display = 'none';
		}else{
			$('morediv').style.display = '';
		}
	}
	//登录按钮
	$('logbtn').onclick = function(){
		window.open('login.asp','_parent','',false);
	}
	//正式注册
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
						alert('注册成功,');
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