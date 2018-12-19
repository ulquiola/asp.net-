// JavaScript Document
function $(id){
	return document.getElementById(id);
}
window.onload = function(){
	showval();
	$('lgname').focus();
	
	$('lgname').onkeydown = function(){
		if(event.keyCode == 13){
			$('lgpwd').select();
		}
	}
	$('lgpwd').onkeydown = function(){
		if(event.keyCode == 13){
			$('lgchk').select();
		}
	}
	$('lgchk').onkeydown = function(){
		if(event.keyCode == 13){
			 chklg();
		}
	}
	$('lgbtn').onclick = chklg;
	function chklg(){
		$('regimg').style.visibility = 'visible';
		if($('lgname').value.match(/^[a-zA-Z_]\w*$/) == null){
			$('regimg').style.visibility = 'hidden';
			alert('请输入合法名称');
			$('lgname').select();
			return false;
		}
		if($('lgname').value == ''){
			$('regimg').style.visibility = 'hidden';
			alert('请输入用户名!');
			$('lgname').focus();
			
			return false;
		}
		if($('lgpwd').value == ''){
			$('regimg').style.visibility = 'hidden';
			alert('请输入密码!');
			$('lgpwd').focus();
			
			return false;
		}
		if($('lgchk').value == ''){
			$('regimg').style.visibility = 'hidden';
			alert('请输入验证码');
			$('lgchk').select();
			
			return false;
		}
		if($('lgchk').value.toUpperCase() != $('chknm').innerText){
			$('regimg').style.visibility = 'hidden';
			alert('验证码输入错误');
			$('lgchk').select();
			
			return false;
		}
		url = 'login_chk.asp?act='+(Math.random())+'&name='+$('lgname').value+'&pwd='+$('lgpwd').value;
		xmlhttp.open('get',url,true);
		xmlhttp.onreadystatechange = function(){
			if(xmlhttp.readystate == 4){
				if(xmlhttp.status == 200){
					msg = xmlhttp.responseText;
					$('regimg').style.visibility = 'hidden';
					if(msg == '1'){
						alert('登录成功');
						location='main.asp';
					}else if(msg == '2'){
						alert('用户名或密码错误');
						$('lgname').select();
					}else if(msg == '3' || msg == '4'){
						alert('非法操作，您的帐号已经被冻结');
					}else{
						alert(msg);
					}
				}
			}
		}
		xmlhttp.send(null);
	}
	$('changea').onclick = showval;
	function showval(){
		xmlhttp.open('get','chkcode.asp',true);
		xmlhttp.onreadystatechange = function(){
			if(xmlhttp.readyState == 4 && xmlhttp.status == 200){
				$('chknm').innerHTML = xmlhttp.responseText;
			}
		}
		xmlhttp.send(null);
	}
	$('fdbtn').onclick = function(){
		fd  = window.open('found.asp','found','width=300,height=200');
		fd.moveTo(screen.width/2,200);
	}
	$('rgbtn').onclick = function(){
		open('register.asp','_parent','',false);
	}
}