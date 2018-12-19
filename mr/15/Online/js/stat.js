// JavaScript Document


function gv_cnzz(of){
	var es = document.cookie.indexOf(";",of);
	if(es==-1) es=document.cookie.length;
	return unescape(document.cookie.substring(of,es));
}
function gc_cnzz(n){
	var arg=n+"=";
	var alen=arg.length;
	var clen=document.cookie.length;
	var i=0;
	while(i<clen){
	var j=i+alen;
	if(document.cookie.substring(i,j)==arg) return gv_cnzz(j);
	i=document.cookie.indexOf(" ",i)+1;
	if(i==0)	break;
	}
	return -1;
}
var ed=new Date();
var now=parseInt(ed.getTime());
var agt=detectOS(); 
var data='l='+escape(location.href)+'&r='+escape(document.referrer)+'&aN='+escape(agt)+'&b='+escape(getOs())+'&s='+window.screen.width+'*'+window.screen.height;


document.write('<img src="http://localhost/visit.asp?'+data+'" border=0 width=10 height=10 id="m">');
alert('http://localhost/visit.asp?'+data);
setInterval("ReLoadImg()",1000*60*10)
function ReLoadImg()
{
	document.getElementById("m").src='http://localhost/visit.asp?'+data;
}
//获取操作系统
function detectOS(){
    var sUserAgent = navigator.userAgent;
    var isWin = (navigator.platform == "Win32") || (navigator.platform == "Windows");
    if(isWin) return "windows";
	var isMac = (navigator.platform == "Mac68K") || (navigator.platform == "MacPPC") || (navigator.platform == "Macintosh");
	if(isMac) return "Mac";
    var isUnix = (navigator.platform == "X11") && !isWin && !isMac;
	if(isUnix) return "Unix";
	
	
    /*if(isWin)
    {
		var isWin95 = sUserAgent.indexOf("Win95") > -1 || sUserAgent.indexOf("Windows 95") > -1;
		if(isWin95) return "Win95";
		var isWin98 = sUserAgent.indexOf("Win98") > -1 || sUserAgent.indexOf("Windows 98") > -1;
		if(isWin98) return "Win98";
		var isWinME = sUserAgent.indexOf("Windows 9x 4.90") > -1 || sUserAgent.indexOf("Windows ME") > -1;
		if(isWinME) return "WinME";
		var isWin2K = sUserAgent.indexOf("Windows NT 5.0") > -1 || sUserAgent.indexOf("Windows 2000") > -1;
		if(isWin2K) return "Win2000";
		var isWinXP = sUserAgent.indexOf("Windows NT 5.1") > -1 || sUserAgent.indexOf("Windows XP") > -1;
		if(isWinXP) return "WinXP";
    }*/
    return "未知";
}

//获取浏览器版本
function getOs() 
{ 
	var Browser_Agent=navigator.userAgent; 
    var OsObject = ""; 
   if(navigator.userAgent.indexOf("MSIE")>0) { 
   		var Version_Start=Browser_Agent.indexOf("MSIE");
   		var Version_End=Browser_Agent.indexOf(";",Version_Start); 
        Actual_Version=Browser_Agent.substring(Version_Start+5,Version_End) 
        return "MSIE+"+Actual_Version; 
   } 
   if(isFirefox=navigator.userAgent.indexOf("Firefox")>0){ 
        return "Firefox"; 
   }
   if(isSafari=navigator.userAgent.indexOf("Safari")>0) { 
        return "Safari"; 
   }  
   if(isCamino=navigator.userAgent.indexOf("Camino")>0){ 
        return "Camino"; 
   } 
   if(isMozilla=navigator.userAgent.indexOf("Gecko/")>0){ 
        return "Gecko"; 
   }
   return "未知";
   
} 

