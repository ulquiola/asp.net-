// ����Ԫ���Ƿ�Ϊ��
function Check(Form)
{
		for(i=0;i<Form.length;i++)
		{  
				if(Form.elements[i].value == "")         //Form������elements������eҪСд
				{
					alert(Form.elements[i].name + "����Ϊ��!");
					return false;
				}
		}
}