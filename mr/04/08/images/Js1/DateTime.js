// 显示日期和时间-------NQ
function ShowTime(Elements)
{
		var temp; 
		var datetime = new Date();
		var year = datetime.getYear();
		var month = datetime.getMonth() + 1;
		var date = datetime.getDate();
		var day = datetime.getDay();
		temp = year+"年"+month+"月"+date+"日 ";
		switch (day)
		{
			case 0:
				temp = temp+"星期日";
				break;
			case 1:
				temp = temp+"星期一";
				break;
			case 2:
				temp = temp+"星期二";
				break;
			case 3:
				temp = temp+"星期三";
				break;
			case 4:
				temp = temp+"星期四";
				break;
			case 5:
				temp = temp+"星期五";
				break;
			case 6:
				temp = temp+"星期六";
				break;
		}
		Elements.innerHTML = temp;   //这个Elements代表(表格)的id
		window.setTimeout("ShowTime(" + Elements.id + ")",1000)   //这里Elements.name或Elements.id不是Elements
		
}
