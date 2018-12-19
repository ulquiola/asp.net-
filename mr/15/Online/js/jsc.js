// JavaScript Document
var oldObj;
function SeeSearchKeyword(obj)
{
	if(obj==oldObj)
	{
		if(obj.parentNode.parentNode.nextSibling.style.display=="none"){
			s_show(obj);
		}
		else{
			s_hid(obj);
		}
	}
	else
	{
		if(oldObj){
			s_hid(oldObj)
		}
		if(obj.innerText=="查看搜索引擎>>"){
			s_show(obj);
		}
	}
	oldObj=obj;
}
function s_show(obj)
{
	obj.innerText="关闭搜索引擎<<";
	obj.parentNode.parentNode.nextSibling.style.display="block";
}
function s_hid(obj)
{
	obj.innerText="查看搜索引擎>>";
	obj.parentNode.parentNode.nextSibling.style.display="none";
}
function GoUrl(strDate,pageFile)
{
	var url="";
	var d=new Date();  
	if(pageFile.indexOf("?")>0)
		pageFile=pageFile+"&"
	else
		pageFile=pageFile+"?"
	switch(strDate)	
	{
		case "today":
			url=pageFile+"startday="+d.format('yyyy-MM-dd')+"&endday="+d.format('yyyy-MM-dd');
			break;
		case "yesterday":
			d.setDate(d.getDate()-1);//
			url=pageFile+"startday="+d.format('yyyy-MM-dd')+"&endday="+d.format('yyyy-MM-dd');
			break;
		case "month":
			var temp=new Date(""+d.getYear()+"/"+d.getMonth()+"/0");
			var tt=d.getYear()+"-"+(d.getMonth()+1)+"-";
			url=pageFile+"startday="+tt+"1"+"&endday="+tt+temp.getDate();
			break;
		case "30":
			var temp=new Date(""+d.getYear()+"/"+d.getMonth()+"/0");
			var tt=d.format('yyyy-MM-dd');
			d.setDate(d.getDate()-30);
			url=pageFile+"startday="+d.format('yyyy-MM-dd')+"&endday="+tt;
			break;
		case "select":
			
			var startDate=document.getElementById("txtStartdate").value;
			var endDate=document.getElementById("txtEnddate").value;
			d.setDate(d.getDate()-30);
			url=pageFile+"startday="+startDate+"&endday="+endDate;
			break;
	}
	location=url;
}
Date.prototype.format = function(format){
    var o =
    {
        "M+" : this.getMonth()+1, //month
        "d+" : this.getDate(),    //day
        "h+" : this.getHours(),   //hour
        "m+" : this.getMinutes(), //minute
        "s+" : this.getSeconds(), //second
        "q+" : Math.floor((this.getMonth()+3)/3),  //quarter
        "S" : this.getMilliseconds() //millisecond
    }
    if(/(y+)/.test(format))
    format=format.replace(RegExp.$1,(this.getFullYear()+"").substr(4 - RegExp.$1.length));
    for(var k in o)
    if(new RegExp("("+ k +")").test(format))
    format = format.replace(RegExp.$1,RegExp.$1.length==1 ? o[k] : ("00"+ o[k]).substr((""+ o[k]).length));
    return format;
}