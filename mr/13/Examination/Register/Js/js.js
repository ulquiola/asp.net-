 /*<meta http-equiv="Content-Type" content="text/html; charset=gb2312">*/
//û��������������Լ���д��һ�������档
function CheckAll()
{
	var len = self.document.all.item("ChkBox");
	for(i=0;i<len.length;i++)
	{
		self.document.all.item("ChkBox",i).checked
	}
}
//��������Լ�д�ģ������ǰѸ�ѡ��ȫѡ��ѡ��
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
//������������ж��Ƿ�ѡ���˸�ѡ��
function Check(form)
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
	}
	if(confirm("ȷ��Ҫɾ����"))
	{
	form.action = "tre_/tre_adm_DelRegister.asp";
	form.submit();
	}
}






