// JavaScript Document
var xmlhttp = false;
try {
  	xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
} catch (e) {
  	try {
    		xmlhttp = new ActivehObject("Microsoft.XMLHTTP");
  	} catch (e2) {}
}

if (!xmlhttp && typeof XMLHttpRequest != "undefined") {
	try
	{
  		xmlhttp = new XMLHttpRequest();
	}catch(e3){ xmlhttp = false;}
}
