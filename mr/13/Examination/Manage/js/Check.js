// ����Ԫ���Ƿ�Ϊ��
function Check(Form)
{
		for(i=0;i<Form.length;i++)
		{  
				if(Form.elements[i].value == "")         //Form������elements������ĸeҪСд
				{
					alert(Form.elements[i].name + "����Ϊ��!");
					Form.elements[i].focus();
					return false;
				}
				if(Form.elements[i].name == "ѧ��֤��")
				{
					if(IsNum(Form.elements[i].value) == false)
					{
							alert("ѧ��֤��ҪȫΪ���֣�");
							Form.elements[i].focus();
							return false;
					}
					if(Form.elements[i].value.length != 16)
					{
							alert("ѧ��֤����16λ!");
							Form.elements[i].focus();
							return false;
					}
				}
		}
}
//�Ѷ��ǲ�������
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