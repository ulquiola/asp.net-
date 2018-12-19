
ht = false;

bb = false;
function AT(ND) {
	document.all("content").value+=ND
}
function ss(size) {
	if (ht) {
		alert("文字大小标记\n设置文字大小.\n可变范围 1 - 6.\n 1 为最小 6 为最大.\n用法: [size="+size+"]这是 "+size+" 文字[/size]");
	} else if (bb) {
		ATT="[font size="+size+"][/font]";
		AT(ATT);
	} else {                       
		tt=prompt("大小 "+size,"文字"); 
		if (tt!=null) {             
			ATT="[font size="+size+"]"+tt;
			AT(ATT);
			ATT="[/font]";
			AT(ATT);
		}        
	}
}

function bold() {
	if (ht) {
		alert("加粗标记\n使文本加粗.\n用法: [b]这是加粗的文字[/b]");
	} else if (bb) {
		ATT="[b][/b]";
		AT(ATT);
	} else {  
		tt=prompt("文字将被变粗.","请在这里输入要加粗的文字！");     
		if (tt!=null) {           
			ATT="[b]"+tt;
			AT(ATT);
			ATT="[/b]";
			AT(ATT);
		}       
	}
}

function xt() {
	if (ht) {
		alert("斜体标记\n使文本字体变为斜体.\n用法: [i]这是斜体字[/i]");
	} else if (bb) {
		ATT="[i][/i]";
		AT(ATT);
	} else {   
		tt=prompt("文字将变斜体","请在这里输入要倾斜的文字！");     
		if (tt!=null) {           
			ATT="[i]"+tt;
			AT(ATT);
			ATT="[/i]";
			AT(ATT);
		}	        
	}
}

function sc(color) {
	if (ht) {
		alert("颜色标记\n设置文本颜色.  任何颜色名都可以被使用.\n用法: [color="+color+"]颜色要改变为"+color+"的文字[/color]");
	} else if (bb) {
		ATT="[font color="+color+"][/font]";
		AT(ATT);
	} else {  
     	tt=prompt("选择的颜色是: "+color,"请在这里输入要改变颜色的文字！");
		if(tt!=null) {
			ATT="[font color="+color+"]"+tt;
			AT(ATT);        
			ATT="[/font]";
			AT(ATT);
		} 
	}
}

function sf(font) {
 	if (ht){
		alert("字体标记\n给文字设置字体.\n用法: [font="+font+"]改变文字字体为"+font+"[/font]");
	} else if (bb) {
		ATT="[font face="+font+"][/font]";
		AT(ATT);
	} else {                  
		tt=prompt("要设置字体的文字"+font,"请在这里输入要改变字体的文字！");
		if (tt!=null) {             
			ATT="[font face="+font+"]"+tt;
			AT(ATT);
			ATT="[/font]";
			AT(ATT);
		}        
	}  
}
function ul() {
  	if (ht) {
		alert("下划线标记\n给文字加下划线.\n用法: [u]要加下划线的文字[/u]");
	} else if (bb) {
		ATT="[u][/u]";
		AT(ATT);
	} else {  
		tt=prompt("下划线文字.","文字");     
		if (tt!=null) {           
			ATT="[u]"+tt;
			AT(ATT);
			ATT="[/u]";
			AT(ATT);
		}	        
	}
}