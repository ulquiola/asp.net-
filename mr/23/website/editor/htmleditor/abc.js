document.onmouseover = doOver;
document.onmouseout  = doOut;
document.onmousedown = doDown;
document.onmouseup   = doUp;
var loaded = 0;
function doOver() {
	var toEl = getReal(window.event.toElement, "className", "coolButton");
	var fromEl = getReal(window.event.fromElement, "className", "coolButton");
	if (toEl == fromEl) return;
	var el = toEl;
	var cDisabled = el.cDisabled;
	cDisabled = (cDisabled != null); 
	if (el.className == "coolButton")
		el.onselectstart = new Function("return false");
	if ((el.className == "coolButton") && !cDisabled) {
		makeRaised(el);
		makeGray(el,false);
	}
}
function doOut() {
	var toEl = getReal(window.event.toElement, "className", "coolButton");
	var fromEl = getReal(window.event.fromElement, "className", "coolButton");
	if (toEl == fromEl) return;
	var el = fromEl;
	var cDisabled = el.cDisabled;
	cDisabled = (cDisabled != null);
	var cToggle = el.cToggle;
	toggle_disabled = (cToggle != null);
	if (cToggle && el.value) {
		makePressed(el);
		makeGray(el,true);
	}
	else if ((el.className == "coolButton") && !cDisabled) {
		makeFlat(el);
		makeGray(el,true);
	}
}
function doUp() {
	el = getReal(window.event.srcElement, "className", "coolButton");
	
	var cDisabled = el.cDisabled;
	cDisabled = (cDisabled != null);
	
	if ((el.className == "coolButton") && !cDisabled) {
		makeRaised(el);
	}
}
function getReal(el, type, value) {
	temp = el;
	while ((temp != null) && (temp.tagName != "BODY")) {
		if (eval("temp." + type) == value) {
			el = temp;
			return el;
		}
		temp = temp.parentElement;
	}
	return el;
}
function findChildren(el, type, value) {
	var children = el.children;
	var tmp = new Array();
	var j=0;	
	for (var i=0; i<children.length; i++) {
		if (eval("children[i]." + type + "==\"" + value + "\"")) {
			tmp[tmp.length] = children[i];
		}
		tmp = tmp.concat(findChildren(children[i], type, value));
	}	
	return tmp;
}
function disable(el) {
	if (document.readyState != "complete") {
		window.setTimeout("disable(" + el.id + ")", 100);
		return;
	}	
	var cDisabled = el.cDisabled;	
	cDisabled = (cDisabled != null);
	if (!cDisabled) {
		el.cDisabled = true;
		el.innerHTML = '<span style="background: buttonshadow; width: 100%; height: 100%; text-align: center;">' +
						'<span style="filter:Mask(Color=buttonface) DropShadow(Color=buttonhighlight, OffX=1, OffY=1, Positive=0); height: 100%; width: 100%%; text-align: center;">' +
						el.innerHTML + '</span>' + '</span>';
		if (el.onclick != null) {
			el.cDisabled_onclick = el.onclick;
			el.onclick = null;
		}
	}
}
function enable(el) {
	var cDisabled = el.cDisabled;	
	cDisabled = (cDisabled != null);	
	if (cDisabled) {
		el.cDisabled = null;
		el.innerHTML = el.children[0].children[0].innerHTML;
		if (el.cDisabled_onclick != null) {
			el.onclick = el.cDisabled_onclick;
			el.cDisabled_onclick = null;
		}
	}
}
function addToggle(el) {
	var cDisabled = el.cDisabled;	
	cDisabled = (cDisabled != null);	
	var cToggle = el.cToggle;	
	cToggle = (cToggle != null);
	if (!cToggle && !cDisabled) {
		el.cToggle = true;		
		if (el.value == null)
			el.value = 0;		
		if (el.onclick != null)
			el.cToggle_onclick = el.onclick;
		else 
			el.cToggle_onclick = "";
		el.onclick = new Function("toggle(" + el.id +"); " + el.id + ".cToggle_onclick();");
	}
}
function removeToggle(el) {
	var cDisabled = el.cDisabled;	
	cDisabled = (cDisabled != null);	
	var cToggle = el.cToggle;	
	cToggle = (cToggle != null);	
	if (cToggle && !cDisabled) {
		el.cToggle = null;
		if (el.value) {
			toggle(el);
		}
		makeFlat(el);		
		if (el.cToggle_onclick != null) {
			el.onclick = el.cToggle_onclick;
			el.cToggle_onclick = null;
		}
	}
}
function toggle(el) {
	el.value = !el.value;	
	if (el.value)
		el.style.background = "URL(/images/tileback.gif)";
	else
		el.style.backgroundImage = "";
}
function makeFlat(el) {
	with (el.style) {
		background = "";
		border = "1px solid buttonface";
		if ((el.id != "more") && (el.id != "fore"))
			padding      = "1px";
	}
}
function makeRaised(el) {
	with (el.style) {
		borderLeft   = "1px solid buttonhighlight";
		borderRight  = "1px solid buttonshadow";
		borderTop    = "1px solid buttonhighlight";
		borderBottom = "1px solid buttonshadow";
		if ((el.id != "more") && (el.id != "fore"))
			padding      = "1px";
	}
}
function makePressed(el) {
	with (el.style) {
		borderLeft   = "1px solid buttonshadow";
		borderRight  = "1px solid buttonhighlight";
		borderTop    = "1px solid buttonshadow";
		borderBottom = "1px solid buttonhighlight";
		if ((el.id != "more") && (el.id != "fore")){
			paddingTop    = "2px";
			paddingLeft   = "2px";
			paddingBottom = "0px";
			paddingRight  = "0px";
		}
	}
}
function makeGray(el,b) {
}

with( document ) {
	write("<style>");
	write(".coolBar {background: buttonface;border-top: 1px solid buttonhighlight;	border-left: 1px solid buttonhighlight; border-bottom: 1px solid buttonshadow; border-right: 1px solid buttonshadow; padding: 2px; font: menu;}");
	write(".coolButton {border: 1px solid buttonface; padding: 1px; text-align: center; cursor: default;}");
	write("</style>");
}

var activeCSS = "border: 1 solid buttonface; color: windowtext; cursor: text;";
var inactiveCSS	= "border: 1 solid window; cursor: hand; color: red;";
var validTextColor = "windowtext";
var invalidTextColor = "buttonshadow";


// 调用格式命令
function doFormat(what) {
	var eb = document.all.editbar;
	eb._editor.execCommand(what, arguments[1]);
}

function doSelectClick(str, el) {
	var Index = el.selectedIndex;
	if (Index != 0){
		el.selectedIndex = 0;
		if (el.id == "specialtype")
			document.all.editbar._editor.specialtype(el.options[Index].value);
		else
			doFormat(str,el.options[Index].value);
	}
}
function doDown() {
	if (window.event.button > 1)
		showHelp();
	document.frames.myEditor.selectRange();
	el = getReal(window.event.srcElement, "className", "coolButton");
	var cDisabled = el.cDisabled;
	cDisabled = (cDisabled != null);
	if ((el.className == "coolButton") && !cDisabled) {
		makePressed(el)
	}
}


function palette(el) {
	var return_value = web_getCol();
	if (return_value != ""){		
		el.style.backgroundColor = return_value;
	}
}

function web_getCol() {
	if (document.createAttribute) {
		var retCol = parent.document.all.dlgColHelper.ChooseColorDlg();
		retCol = retCol.toString(16);
		if( retCol == 0 )
			return null;
		if( retCol.length < 6 ) {
			var tmp = "000000".substring(0,6-retCol.length);
			retCol = tmp.concat(retCol);
		}
	}else{
		var basCol = showModalDialog("palette.html","","dialogWidth:18;dialogHeight:16;dialogTop:50;dialogLeft:500;status:no;");
		if( basCol != null ) {
			retCol = basCol;
		}
	}

	var Hclr = retCol.charAt(0);
	if( Hclr != "#" )
		retCol = "#" + retCol;
	return retCol;
}


function showHelp(){	
	showModalDialog("help.htm","","dialogWidth:35;dialogHeight:25;dialogTop:100;dialogLeft:200;status:no;");
}

function SaveMyEdit() {
	savefile();
	parent.document.HtmlEditForm.content.value = parent.document.HtmlEditForm.body.value;
	parent.document.HtmlEditForm.submit();
}


function showmore(){
	editbar.style.display = "none";
	editbar1.style.display = "block";
	more.style.display = "none";
	fore.style.display = "block";
}


function showfore(){
	editbar.style.display = "block";
	editbar1.style.display = "none";
	more.style.display = "block";
	fore.style.display = "none";
}


function addTable(){
	var iHeight, iWidth;
	var sTable = prompt("请输入表格的行数和列数，中间用逗号分隔，\n如：3,4", "2,2");
	if (sTable == null){return 0;}
	var i = sTable.indexOf(",");	
	if (i < 1){alert("输入内容中没有找到逗号！");return 0;}
	iHeight = parseInt(sTable.substr(0,i++));
	if (isNaN(iHeight)){alert("表格行数不是整数！");return 0;}
	if  (iHeight < 1){alert("表格行数不是正整数！");return 0;}
	if  (iHeight > 20000){alert("表格行数过大！");return 0;}
	iWidth = parseInt(sTable.substr(i));
	if (isNaN(iWidth)){alert("表格列数不是整数！");return 0;}
	if  (iWidth < 1){alert("表格列数不是正整数！");return 0;}
	if  (iWidth > 200){alert("表格列数过大！");return 0;}
	var sHTML = "<table class=ubb cellspacing=0>";
	for(i=0;i<iHeight;i++){
		sHTML += "<tr>"
		for(j=0;j<iWidth;j++)
			sHTML += "<td class=ubb>&nbsp; </td>";
		sHTML += "</tr>";
	}
	sHTML += "</table>";
	myEditor.pasteHTML(sHTML);
}


function nowinstatus(){
	window.status=''
}


function fixSize() {
	document.all.edit1.style.height = Math.max(document.body.clientHeight - document.all.edit1.offsetTop, 0);
}


function init() {
	fixSize();	
	if (loaded){
		parent.document.all.editor.style.display = "block";
		parent.status = "";		
	}
	else
		loaded = 1;	
}


function AddImg(){	
	document.all.UploadImg.style.display = "block";
}


function CancelAddImg(){
	document.all.UploadImg.style.display = "none";
}


function UploadComplete(URL){
	if( URL.length <4 )
		return false;
	if( (URL != null) && (URL != "") ) {
		var fileExt = "";
		if( URL.indexOf(".") >=0 ) {
			var fileExtArr = URL.split(".");
			fileExt = fileExtArr[fileExtArr.length-1];
			fileExt = fileExt.toLowerCase();
		}
		switch( fileExt ) {
			case "zip" :
			case "rar" :
			case "doc" :
			case "htm" :
			case "html" :
				InsertDownload( URL, fileExt );
				break;
			case ""   :
				InsertDownload( URL, "Unknown" );
				break;
			case "mid" :
				InsertBgSound( URL );
				break;
			case "swf" :
				InsertSwf( URL );
				break;
			default   :
				if( URL.indexOf(":") == -1 ) {
					doFormat("InsertImage", URL);
				}else{
					doFormat("InsertImage", URL);
				}
				break;
		}
	}
	document.all.UploadImg.style.display = "none";
	document.forms["upload"].reset();
}



function InsertDownload( FilePath, Ext ) {

	var str = "<a href=\"" + FilePath + "\" target=_blank>";
	str += "<img src=\"/images/Ext/" + Ext + ".gif\" border=0>";
	str += "</a>";

	myEditor.textEdit.focus();
	var edtiarea = myEditor.textEdit.document.selection.createRange();
	edtiarea.pasteHTML( str );
	myEditor.textEdit.focus();
	edtiarea.select();
}


function InsertBgSound( SoundPath ) {

	var str = "<head><BGSOUND balance=0 src=\"" + SoundPath + "\" volume=-20 loop=infinite></head>";

	myEditor.textEdit.focus();
	var edtiarea = myEditor.textEdit.document.selection.createRange();
	edtiarea.pasteHTML( str );
	myEditor.textEdit.focus();
	edtiarea.select();
}


function InsertSwf( swf ) {

	var swfw = GetCookie("swfw");
	var swfh = GetCookie("swfh");
	var strSwfwh = "width=" + 100 + " height=" + 100 + " ";


	strSwf = "<embed wmode=transparent src=\"" + swf + "\" ";
	strSwf += strSwfwh;
	strSwf += "type=\"application/x-shockwave-flash\" pluginspage=\"http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash\">";
	strSwf += "</embed>\n";

	myEditor.textEdit.focus();
	var edtiarea = myEditor.textEdit.document.selection.createRange();
	edtiarea.pasteHTML( strSwf );
	myEditor.textEdit.focus();
	edtiarea.select();
}

function getCookieVal(offset) {
	var endstr=document.cookie.indexOf(";",offset);
	if( endstr==-1 )
		endstr=document.cookie.length;
	return unescape(document.cookie.substring(offset,endstr));
}

function GetCookie(name) {
	var arg = name + "=";
	var alen = arg.length;
	var clen = document.cookie.length;
	var i = 0;
	var j;
	while( i<clen ) {
		j = i + alen;
		if( document.cookie.substring(i,j)==arg )
			return getCookieVal(j);
		i = document.cookie.indexOf(" ",i)+1;
		if(i==0)
			break;
	}
	return null;
}


function savefile(){
	if (edit1.format == "ABC")
		parent.document.HtmlEditForm.body.value = myEditor.textEdit.document.body.innerText;
	else
		parent.document.HtmlEditForm.body.value = myEditor.textEdit.document.body.innerHTML;
}


