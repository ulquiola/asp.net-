var format = "HTML";
var initHTML = "";
var edit;
var RangeType;
function setFocus() {
	textEdit.focus();
}
function selectRange(){
	edit = textEdit.document.selection.createRange();
	RangeType = textEdit.document.selection.type;
}
function execCommand(command) {	
	if (format == "HTML"){
		setFocus();
		selectRange();	
		if ((command == "Undo") || (command == "Redo"))
			document.execCommand(command);
		else{
			if (arguments[1]==null)				
				edit.execCommand(command);
			else
				edit.execCommand(command, false, arguments[1]);}
		textEdit.focus();
		if (RangeType != "Control") edit.select();
	}	
}
function swapModes(Mode) {	
	switch(Mode){
		case 1:
			if (format == "ABC"){
				textEdit.document.body.innerHTML = textEdit.document.body.innerText;
				textEdit.document.body.style.fontFamily = "";
				textEdit.document.body.style.fontSize ="";
			}
			else{
				initHTML = textEdit.document.body.innerHTML;
				initEditor();
			}
			format = "HTML";
			break;	
		case 2:
			if (format == "PREVIEW"){
				initHTML = textEdit.document.body.innerHTML;
				initEditor();
			}	
			textEdit.document.body.innerText = textEdit.document.body.innerHTML;
			textEdit.document.body.style.fontFamily = "Fixedsys";
			textEdit.document.body.style.fontSize = "9pt";
			format = "ABC";
			break;
		case 3:
			var strHTML = "";
			if(format == "ABC"){
				strHTML = textEdit.document.body.innerText;
				textEdit.document.body.style.fontFamily = "";
				textEdit.document.body.style.fontSize ="";				
			}
			else{
				strHTML = textEdit.document.body.innerHTML;				
			}			
			format = "PREVIEW";
			textEdit.document.designMode="Off";
			textEdit.document.open();
			textEdit.document.write(strHTML);
			textEdit.document.close();
			if(textEdit.document.styleSheets.length == 0){
				textEdit.document.createStyleSheet();
				var oSS = textEdit.document.styleSheets[0];
				oSS.addRule("TABLE.ubb","border: 1px solid #A9A9A9;FONT-SIZE: 9pt; ");
				oSS.addRule("TD.ubb","border: 1px solid #A9A9A9;FONT-SIZE: 9pt; ");
				oSS.addRule("BODY","FONT-SIZE: 9pt;");
				oSS.addRule("IMG","border: 0");
			}
			break;
		default:
			return(0);
	}
	textEdit.focus();
	return(1);
}
function specialtype(Mark){
	if (format == "HTML"){
		var strHTML;
		setFocus();
		selectRange();	
		if (RangeType == "Text"){
			strHTML = "<" + Mark + ">" + edit.text + "</" + Mark + ">"; 
			edit.pasteHTML(strHTML);
			textEdit.focus();
			edit.select();			
		}
	}
}
function pasteHTML(HTML){	
	if (format == "HTML"){
		setFocus();
		selectRange();
		edit.pasteHTML(HTML);
		textEdit.focus();
		if (RangeType != "Control") edit.select();
	}
}
function initEditor() {
	textEdit.document.designMode="On";
	textEdit.document.open();
	textEdit.document.write(initHTML);
	textEdit.document.close();
	initHTML = "";
	if(textEdit.document.styleSheets.length == 0){
		textEdit.document.createStyleSheet();
		var oSS = textEdit.document.styleSheets[0];
		oSS.addRule("TABLE.ubb","border: 1px solid #A9A9A9;FONT-SIZE: 9pt; ");
		oSS.addRule("TD.ubb","border: 1px solid #A9A9A9;FONT-SIZE: 9pt; ");
		oSS.addRule("IMG","border: 0");
		oSS.addRule("BODY","font-size: 9pt");
	}	
}
function init() {
	initEditor();
	with (parent){
		if (loaded){
			parent.status = "";
		}
		else
			loaded = 1;	
	}
}
initHTML = parent.parent.strHTML;
window.onload = init