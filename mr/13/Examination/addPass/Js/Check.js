// 检查表单元素是否为空
function Check(Form)
{
		for(i=0;i<Form.length;i++)
		{  
				if(Form.elements[i].value == "")         //Form的属性elements的首字e要小写
				{
					alert(Form.elements[i].name + "不能为空!");
					return false;
				}
		}
}