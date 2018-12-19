 /*<meta http-equiv="Content-Type" content="text/html; charset=gb2312">*/
//没用这个函数，我自己又写了一个在下面。
function CheckAll()
{
	var len = self.document.all.item("ChkBox");
	for(i=0;i<len.length;i++)
	{
		self.document.all.item("ChkBox",i).checked
	}
}
//这个是我自己写的，作用是把复选框全选或反选。
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
//这个的作用是判断是否选择了复选框。
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
			alert("请选择要删除的项目");
			return false;
		}
	}
	else
	{
		if(len.checked == false)
		{
			alert("请选择要删除的项目");
			return false;
		}
	}
	if(confirm("确定要删除吗？"))
	{
	form.action = "tre_/tre_adm_DelRegister.asp";
	form.submit();
	}
}






