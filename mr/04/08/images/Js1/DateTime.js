// ��ʾ���ں�ʱ��-------NQ
function ShowTime(Elements)
{
		var temp; 
		var datetime = new Date();
		var year = datetime.getYear();
		var month = datetime.getMonth() + 1;
		var date = datetime.getDate();
		var day = datetime.getDay();
		temp = year+"��"+month+"��"+date+"�� ";
		switch (day)
		{
			case 0:
				temp = temp+"������";
				break;
			case 1:
				temp = temp+"����һ";
				break;
			case 2:
				temp = temp+"���ڶ�";
				break;
			case 3:
				temp = temp+"������";
				break;
			case 4:
				temp = temp+"������";
				break;
			case 5:
				temp = temp+"������";
				break;
			case 6:
				temp = temp+"������";
				break;
		}
		Elements.innerHTML = temp;   //���Elements����(���)��id
		window.setTimeout("ShowTime(" + Elements.id + ")",1000)   //����Elements.name��Elements.id����Elements
		
}
