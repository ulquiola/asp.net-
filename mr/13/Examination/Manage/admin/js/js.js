//����CheckAll�������ǽ���ѡ��ȫѡ��ѡ
function CheckAll(elementsA,elementsB)
{
	var len = elementsA;
	if(len.length > 0)
	{
		for(i=0;i<len.length;i++)
		{
			elementsA[i].checked = true;
		}
		if(elementsB.checked ==false)
		{
			for(j=0;j<len.length;j++)
			{
				elementsA[j].checked = false;
			}
		}
	}
	else
	{
		len.checked = true;
		if(elementsB.checked == false)
		{
			len.checked = false;
		}
	}
}
//����Check���������ж��Ƿ�ѡ���˸�ѡ��
function Check(form,id)
{
	var flag = false;
	var len = self.document.all.item("ChkBox");
	var path ="";
	if(len.length > 0)
	{
		for(j=0;j<len.length;j++)
		{
			if(self.document.all.item("ChkBox",j).checked)
			{
				if(self.document.all.item("ChkBox",j).value == id){
					alert('������ɾ����ǰ����Ա!');
					self.document.all.item("ChkBox",j).checked = false;
					return false;
				}
				flag = true;
			}
		}
		if(!flag)
		{
			alert("��ѡ��Ҫɾ������Ŀ");
			return false;
		}
	}
	else
	{
		if(len.checked == false)
		{
			alert("��ѡ��Ҫɾ������Ŀ");
			return false;
		}
		if(self.document.all.item("ChkBox").value == id){
			alert('������ɾ����ǰ�û�');
			self.document.all.item("ChkBox").checked = false;
		}
		return false;
	}
	if(confirm("ȷ��Ҫɾ����"))
	{
	form.action = "tre_/tre_adm_DelAdmin.asp";
	form.submit();
	}
}


