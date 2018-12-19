 /*<meta http-equiv="Content-Type" content="text/html; charset=gb2312">*/
//作用是把复选框全选或反选。
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
	var frname = form.name; //这里浪费了一些时间，因为form是表单对象，所以要取它的名子要用name属性。
	var num = frname.indexOf("result");
	if( num > 0)
	{
		path = "tre_/tre_adm_Delresult.asp";
	}
	else
	{
		path = "tre_/tre_adm_DelQuestion.asp";
	}
	if(confirm("确定要删除吗？"))
	{
	form.action = path;
	form.submit();
	}
}
//隐藏表格。
function hid_td(td)
{
	td.style.display = "none";
}
//显示表格。
function show_td()
{
	td_fra.style.display = "";
}	
//下拉列表框联动
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
//打开新窗口
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
//修改试题
function upques_danorduo(name,id,proid,lesid,taotiid)
{
	location.href = "adm_UpdateQuestion.asp?name="+name+"&id="+id+"&proid="+proid+"&lesid="+lesid+"&taotiid="+taotiid;
}
//修改试题之重置
function rel_load(id,proid,lesid,taotiid)
{
	location.href = "adm_UpdateQuestion.asp?id="+id+"&proid="+proid+"&lesid="+lesid+"&taotiid="+taotiid;
}







