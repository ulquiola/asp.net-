<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn/conn.asp"-->
<%
if request.form("sousuo")<>"" then
	sousuo=request.form("sousuo")
else
	sousuo=request.querystring("sousuo")
end if 
set rs=server.CreateObject("adodb.recordset")
rs1="select * from tb_sousuo where content like '%"&sousuo&"%'"
rs.open rs1,conn,1,1
%>
<script language="javascript">
	function clockon(){
		var now=new Date();
		var year=now.getYear();
		var month=now.getMonth();
		var date=now.getDate();
		var day=now.getDay();
		var hour=now.getHours();
		var minu=now.getMinutes();
		var sec=now.getSeconds();
		var week;
		month=month+1;
		if(month<10) month="0"+month;
		if(date<10) date="0"+date;
		if(hour<10) hour="0"+hour;
		if(minu<10) minu="0"+minu;
		if(sec<10) sec="0"+sec;
		switch (day)
		  { case 1:
		        week="星期一";
				break;
			case 2:
		        week="星期二";
				break;
			case 3:
		        week="星期三";
				break;
			case 4:
		        week="星期四";
				break;
			case 5:
		        week="星期五";
				break;
			case 6:
		        week="星期六";
				break;
			default:
		        week="星期日";
				break;
		  }
		var time="";
		time=year+"年"+month+"月"+date+"日 "+week+" "+hour+":"+minu+":"+sec;
		if(document.all){
			bgclock.innerHTML=time;
		}
		var timer=setTimeout("clockon()",200);
	}  
</script>
<html>
<head>
<script language="javascript">
<!-- Hide this script from old browsers --
var speed = 50
var pause = 1500
var timerID = null
var bannerRunning = false
var ar = new Array()
ar[0] = "欢迎来到明日科技有限公司！"
ar[1] = "欢迎光临明日在线搜索网！！"
ar[2] = "请您记住本站域名 吉林省明日科技有限公司 http://www.mingrisoft.COM 谢谢光临!!"
var message = 0
var state = ""
clearState()
function stopBanner() {
        if (bannerRunning)
                clearTimeout(timerID)
        bannerRunning = false
}

function startBanner() {
        stopBanner()
        showBanner()
}
function clearState() {
        state = ""
        for (var i = 0; i < ar[message].length; ++i) {
                state += "0"
        }
}
function showBanner() {
        if (getString()) {
                message++
                if (ar.length <= message)
                        message = 0
                clearState()
                timerID = setTimeout("showBanner()", pause)
                bannerRunning = true
        } else {
                var str = ""
                for (var j = 0; j < state.length; ++j) {
                        str += (state.charAt(j) == "1") ? ar[message].charAt(j) : "     "
                }
                window.status = str
                timerID = setTimeout("showBanner()", speed)
                bannerRunning = true
        }
}
function getString() {
        var full = true
        for (var j = 0; j < state.length; ++j) {
                if (state.charAt(j) == 0)
                        full = false
        }
        if (full)
                return true
        while (1) {
                var num = getRandom(ar[message].length)
                if (state.charAt(num) == "0")
                        break
        }
        state = state.substring(0, num) + "1" + state.substring(num + 1, state.length)
        return false
}
function getRandom(max) {
        return Math.round((max - 1) * Math.random())
}
// -- End Hiding Here -->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>明日在线搜索</title>
<style>
td{
font-size:9pt;
}
</style>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
<style type="text/css">
<!--
.style4 {font-size: 16px}
.style22 {font-size: 14pt}
.style26 {font-size: 12px}
-->
</style>
<link href="style1.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	background-color: #bdcff7;
	background-image: url(images/bg.gif);
	margin-top: 0px;
}
.STYLE28 {font-size: 13px}
-->
</style></head>
<body onLoad="clockon(),startBanner()">
<table width="779" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="186" colspan="3" valign="top" bgcolor="#FFFFFF">
      <!--#include file="shangbiao.asp"-->
      <table width="780" height="34" border="0" align="center" cellpadding="0" cellspacing="0"><form name="form1" method="post" action="sousuo.asp">   
                <tr>
                  <td width="156" height="53" align="center" background="images/sousou1.gif">                  　</td>
                  <td width="117" align="center" background="images/sousou2.gif">请输入搜索关键字：</td>
                  <td width="235" align="center" background="images/sousou2.gif"><div align="left">
                    <input name="sousuo" type="text" id="sousuo" style="height:25px" size="32">
                  </div></td>
                  <td width="60" align="center" background="images/sousou2.gif"><input name="imageField" type="image" src="images/souan1.gif" width="42" height="20" border="0"></td>
                  <td width="92" height="53" align="center" background="images/sousou2.gif"><div align="left"><a href="gaojisousuo.asp"><img src="images/souan2.gif" width="77" height="20" border="0"></a></div></td>
                  <td width="120" height="53" align="center" background="images/sousouhou.gif"><a href="sousuo1.asp"></a></td>
                </tr>
                <tr>
              <td colspan="6"></td>
            </tr>
</form>			
      </table>
	  <table width="567" height="350" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="111" align="center" valign="middle" class="style4"><strong>针对性搜索</strong></td>
        </tr>
        <tr>
          <td align="center" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="25%" height="35" align="center" class="STYLE28"><a href="personnet.asp">人才网站</a></td>
              <td width="25%" height="35" align="center" class="STYLE28"><a href="uselanguage.asp">常用术语</a></td>
              <td width="25%" height="35" align="center" class="STYLE28"><a href="booksdata.asp">图书资源</a></td>
              <td width="25%" height="35" align="center" class="STYLE28"><a href="ITman.asp">IT人物简介</a></td>
            </tr>
            <tr>
              <td width="25%" height="35" align="center" class="STYLE28"><a href="programme.asp">编程网站</a></td>
              <td width="25%" height="35" align="center" class="STYLE28"><a href="openlanguage.asp">开发语言</a></td>
              <td width="25%" height="35" align="center" class="STYLE28"><a href="story.asp">创业故事</a></td>
              <td width="25%" height="35" align="center" class="STYLE28"><a href="ITstory.asp">IT企业故事</a></td>
            </tr>
          </table></td>
        </tr>
    </table>    </td>
  </tr>
  <tr>
    <td height="22" colspan="3" bgcolor="#FFFFFF">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#FF00CC" bordercolordark="#FFFFFF">
<tr>
<td height="20" align="center"><table width="101%" border="0" align="center" cellPadding="0" cellSpacing="0" bgcolor="#CCCCCC">
<script type="text/javascript">
function search4()
{
if(websearch.google.checked)
window.open("http://www.google.cn/custom?&sa=%CB%D1%CB%F7&client=pub-2484478384582955&forid=1&ie=GB2312&oe=GB2312&cof=GALT%3A%23008000%3BGL%3A1%3BDIV%3A%23336699%3BVLC%3A663399%3BAH%3Acenter%3BBGC%3AFFFFFF%3BLBGC%3A336699%3BALC%3A0000FF%3BLC%3A0000FF%3BT%3A000000%3BGFNT%3A0000FF%3BGIMP%3A0000FF%3BLH%3A31%3BLW%3A88%3BL%3Ahttp%3A%2F%2Fwww.910job.net%2Fimg%2Flogo.gif%3BS%3Ahttp%3A%2F%2Fwww.910job.net%3BFORID%3A1%3B&hl=zh-CN&q="+websearch.key.value,"mspg6");
if(websearch.baidu.checked)
window.open("http://www.baidu.com/s?tn=teacherjob&ct=&lm=&z=&rn=&word="+websearch.key.value,"mspg9");
if(websearch.sina.checked)
window.open("http://isou.com/s?pid=33640&fromlianmeng&k="+websearch.key.value,"mspg0");
if(websearch.zhsou.checked)
window.open("http://p.zhongsou.com/p?k=teacherjobs&w="+websearch.key.value,"mspg4");
if(websearch.sohu.checked)
window.open("http://so.sohu.com/web?pid=37021com&query="+websearch.key.value,"mspg1");
if(websearch.yahoo.checked)
window.open("http://business.cn.yahoo.com/bso?ei=gbk&pid=un_63495_239_13_0&f=C63495_5_0-239&p="+websearch.key.value,"mspg2");
if(websearch.yeah.checked)
window.open("http://cha.so.163.com/so.php?q="+websearch.key.value,"mspg3");
if(websearch.tom.checked)
window.open("http://search.tom.com/search.php?&word="+websearch.key.value,"mspg7");
if(websearch.MSN.checked)
window.open("http://beta.search.msn.com.cn/results.aspx?FORM=MSNH&cp=936&q="+websearch.key.value,"mspg8");
if(websearch.sogou.checked)
window.open("http://www.sogou.com/websearch/corp/search.jsp?&searchtype=1&pid=teacherjobs&class=1&cpc=SOGOU&query="+websearch.key.value,"mspg10");
if(websearch.soso.checked)
window.open("http://www.soso.com/q?w="+websearch.key.value,"mspg10");
if(websearch.tianwang.checked)
window.open("http://www.tianwang.com/cgi-bin/tw?&range=0&cd=03&submit.x=26&submit.y=12&word="+websearch.key.value,"mspg10");
return false; 
}
</script>
<form name="websearch" onSubmit="return(search4())">
<td height="32" align="center">
<input onBlur="if (value ==''){value='请输入关键字'}" onMouseOver="this.focus()" style="FONT-SIZE: 12px" onFocus="this.select()" onClick="if(this.value=='请输入关键字')this.value=''" size="14" name="key">
<input type="submit" value="搜索" name="submit" /><input name="baidu" type="checkbox" value="baidu" checked />百度<input type="checkbox" value="sina" name="sina" />新浪<input name="sohu" type="checkbox" value="sohu" />搜狐<input type="checkbox" value="sogou" name="sogou" />搜狗<input type="checkbox" value="yahoo" name="yahoo" />雅虎<input type="checkbox" value="yeah" name="yeah" />网易<input type="checkbox" value="google" name="google" />google<input name="zhsou" type="checkbox" value="zhsou" />中搜<input name="tom" type="checkbox" value="tom" />tom<input name="MSN" type="checkbox" value="MSN" />MSN<input name="soso" type="checkbox" value="soso" />腾讯<input name="tianwang" type="checkbox" id="tianwang" value="tianwang" />天网</td>
<td width="2"></form>
</tr>
</table></td>
</tr>
</table>
	</td>
  </tr></tr>
</tr>

  <tr>
    <td width="192" height="50" valign="top" background="images/di.gif" bgcolor="#FFFFFF">&nbsp;</td>
    <td width="392" height="50" align="center" valign="middle" background="images/di.gif" bgcolor="#FFFFFF" class="STYLE5">All CopyRight(C) reserved 2008 吉林省明日科技有限公司<br>
      <br>
    服务热线：0431-84978981 84978982 </td>
    <td width="196" height="50" valign="top" background="images/di.gif" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
</table>
</body>
</html>
