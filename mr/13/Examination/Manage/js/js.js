 /*<meta http-equiv="Content-Type" content="text/html; charset=gb2312">*/
//�����ǰѸ�ѡ��ȫѡ��ѡ��
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
	var frname = form.name; //�����˷���һЩʱ�䣬��Ϊform�Ǳ���������Ҫȡ��������Ҫ��name���ԡ�
	var num = frname.indexOf("result");
	if( num > 0)
	{
		path = "tre_/tre_adm_Delresult.asp";
	}
	else
	{
		path = "tre_/tre_adm_DelQuestion.asp";
	}
	if(confirm("ȷ��Ҫɾ����"))
	{
	form.action = path;
	form.submit();
	}
}
//���ر��
function hid_td(td)
{
	td.style.display = "none";
}
//��ʾ���
function show_td()
{
	td_fra.style.display = "";
}	
//�����б������
function localhost(form)
{	
	form.target="mainr";
	form.action = "adm_Mainleft_sel.asp";
	form.submit();
}
function localhost2(form,tag)
{
	form.target="mainr";
	form.action = "adm_Mainleft_sel.asp?flag=" + tag;
	form.submit();
	//location.href = "adm_Mainleft_sel.asp?proid="+proval+"&lessonid="+lessonval;
}
function goto(form)
{
	form.target="mainr";
	form.action = "questions/adm_AddQuestion.asp";
	form.submit();
}
//���´���
function openwin(url)
{
	showModalDialog(url,"","help=no;status=no;")
}
//
function danorduo(name,proid,lesid,taotiid)
{
	location.href = "adm_AddQuestion.asp?name="+name+"&proid="+proid+
	"&lessonid="+lesid+"&sel_taoti="+taotiid;
}
//�޸�����
function upques_danorduo(name,id,proid,lesid,taotiid)
{
	location.href = "adm_UpdateQuestion.asp?name="+name+"&id="+id+"&proid="+proid+"&lesid="+lesid+"&taotiid="+taotiid;
}
//�޸�����֮����
function rel_load(id,proid,lesid,taotiid)
{
	location.href = "adm_UpdateQuestion.asp?id="+id+"&proid="+proid+"&lesid="+lesid+"&taotiid="+taotiid;
}







