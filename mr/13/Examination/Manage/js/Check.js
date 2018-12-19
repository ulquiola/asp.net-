// 检查表单元素是否为空
function Check(Form)
{
		for(i=0;i<Form.length;i++)
		{  
				if(Form.elements[i].value == "")         //Form的属性elements的首字母e要小写
				{
					alert(Form.elements[i].name + "不能为空!");
					Form.elements[i].focus();
					return false;
				}
				if(Form.elements[i].name == "学生证号")
				{
					if(IsNum(Form.elements[i].value) == false)
					{
							alert("学生证号要全为数字！");
							Form.elements[i].focus();
							return false;
					}
					if(Form.elements[i].value.length != 16)
					{
							alert("学生证号是16位!");
							Form.elements[i].focus();
							return false;
					}
				}
		}
}
//叛断是不是数字
function IsNum(Value)
{
		var NumStr = "0123456789";
		for(j=0;j<Value.length;j++)
		{
				if(NumStr.indexOf(Value.substr(j,1)) < 0)
				{
						return false;
				}
		}
}