//函数CheckAll的作用是将复选框全选或反选
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
//函数Check的作用是判断是否选择了复选框
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
					alert('不允许删除当前管理员!');
					self.document.all.item("ChkBox",j).checked = false;
					return false;
				}
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
		if(self.document.all.item("ChkBox").value == id){
			alert('不允许删除当前用户');
			self.document.all.item("ChkBox").checked = false;
		}
		return false;
	}
	if(confirm("确定要删除吗？"))
	{
	form.action = "tre_/tre_adm_DelAdmin.asp";
	form.submit();
	}
}


