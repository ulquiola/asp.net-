
ht = false;

bb = false;
function AT(ND) {
	document.all("content").value+=ND
}
function ss(size) {
	if (ht) {
		alert("���ִ�С���\n�������ִ�С.\n�ɱ䷶Χ 1 - 6.\n 1 Ϊ��С 6 Ϊ���.\n�÷�: [size="+size+"]���� "+size+" ����[/size]");
	} else if (bb) {
		ATT="[font size="+size+"][/font]";
		AT(ATT);
	} else {                       
		tt=prompt("��С "+size,"����"); 
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
		alert("�Ӵֱ��\nʹ�ı��Ӵ�.\n�÷�: [b]���ǼӴֵ�����[/b]");
	} else if (bb) {
		ATT="[b][/b]";
		AT(ATT);
	} else {  
		tt=prompt("���ֽ������.","������������Ҫ�Ӵֵ����֣�");     
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
		alert("б����\nʹ�ı������Ϊб��.\n�÷�: [i]����б����[/i]");
	} else if (bb) {
		ATT="[i][/i]";
		AT(ATT);
	} else {   
		tt=prompt("���ֽ���б��","������������Ҫ��б�����֣�");     
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
		alert("��ɫ���\n�����ı���ɫ.  �κ���ɫ�������Ա�ʹ��.\n�÷�: [color="+color+"]��ɫҪ�ı�Ϊ"+color+"������[/color]");
	} else if (bb) {
		ATT="[font color="+color+"][/font]";
		AT(ATT);
	} else {  
     	tt=prompt("ѡ�����ɫ��: "+color,"������������Ҫ�ı���ɫ�����֣�");
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
		alert("������\n��������������.\n�÷�: [font="+font+"]�ı���������Ϊ"+font+"[/font]");
	} else if (bb) {
		ATT="[font face="+font+"][/font]";
		AT(ATT);
	} else {                  
		tt=prompt("Ҫ�������������"+font,"������������Ҫ�ı���������֣�");
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
		alert("�»��߱��\n�����ּ��»���.\n�÷�: [u]Ҫ���»��ߵ�����[/u]");
	} else if (bb) {
		ATT="[u][/u]";
		AT(ATT);
	} else {  
		tt=prompt("�»�������.","����");     
		if (tt!=null) {           
			ATT="[u]"+tt;
			AT(ATT);
			ATT="[/u]";
			AT(ATT);
		}	        
	}
}