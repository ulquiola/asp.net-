<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨����</title>
<!--#include file="../include/conn.asp" -->
<!--#include file="../include/include.asp" -->
<!--#include file="chk_admin.asp" -->

<script language = "JavaScript">
<!-- �˴��������Ҫ������Ϊ�����û���ѡ����Ʒ�����ͬʱ��ʾ��Ӧ�ķ��� -->
var onecount;
onecount=0;
subcat = new Array();	//�������飬Ŀ���ǰ�һ�����ɴ�����з���
<%
i=0
sql="select * from [class] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof			''��ѯ����ѭ��������з���
	i=i+1	''���������±꣬���ʱASP���� i ��ֵΪ 1���������±��ʼֵΪ 0����������� i-1 ����Ϊ�˷��������±�Ĺ��� 
			''���ݿ����ж��ٷ��ϵ����ݣ���������������ж��٣���������ֵ���Ǳ������������ 1����Ϊ�����±��� 0 ��ͷ�� 
%>
	subcat[<%=i-1%>] = new Array("<%=rs("mingcheng")%>","<%=rs("bigclassid")%>","<%=rs("id")%>");
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
onecount=<%response.Write i ''����������������ȻASP���� i ��ѭ�����⣬�� i ��ѭ�������Ѿ���������ֵ%>;

function changelocation(locationid)
{
	document.myform.classid.length = 0;
	var locationid=locationid;	//��ʼ�� locationid ����������ֵȡ�õĴ��� ID
	var i;
	for (i=0;i < onecount; i++)		//����������������������������ѭ��
	{
		if (subcat[i][1] == locationid)	//�����з������ж϶�Ӧ�Ĵ��� ID
		{ 
			 document.myform.classid.options[document.myform.classid.length] = new Option(subcat[i][0], subcat[i][2]);
			 //�����ѡ����Ʒ�����ͬʱ��ʾ��Ӧ�ķ���ĸ�ֵ
		}        
	}
}    
<!-- end -->
</script>

<script language = "JavaScript">
function addpro()
{
	if(document.myform.jianjie.value=="") 
	{
		document.myform.jianjie.focus();
		alert("��������Ʒ��飡");
		return false;
	}
	if(document.myform.riqi.value=="") 
	{
		document.myform.riqi.focus();
		alert("������������ڣ�");
		return false;
	}
	if(document.myform.mingcheng.value=="") 
	{
		document.myform.mingcheng.focus();
		alert("��������Ʒ���ƣ�");
		return false;
	}
	if(document.myform.shichang.value=="") 
	{
		document.myform.shichang.focus();
		alert("�������г��۸�");
		return false;
	}
	if(document.myform.huiyuan.value=="") 
	{
		document.myform.huiyuan.focus();
		alert("�������Ա�۸�");
		return false;
	}
	if(document.myform.xinghao.value=="") 
	{
		document.myform.xinghao.focus();
		alert("��������Ʒ�ͺţ�");
		return false;
	}
	if(document.myform.shuoming.value=="") 
	{
		document.myform.shuoming.focus();
		alert("��������Ʒ˵����");
		return false;
	}
	if(document.myform.beizhu.value=="") 
	{
		document.myform.beizhu.focus();
		alert("��������Ʒ��ע��");
		return false;
	}
	if(document.myform.shuliang.value=="") 
	{
		document.myform.shuliang.focus();
		alert("��������Ʒ������");
		return false;
	}
}
</script>
</head>
<body>
<%
if request("action")="add" then	'�����ر�����ȡ action ֵ���ж��Ƿ��ɸñ��ύ������
	if request("jianjie")="" or request("riqi")="" or request("mingcheng")="" or request("shichang")="" or request("huiyuan")="" or request("xinghao")="" or request("dengji")="" or session("tupian")="" or request("shuoming")="" or request("beizhu")="" or request("bigclassid")="" or request("classid")="" or request("shuliang")="" then
		response.Write("<script>alert('����ϸ��д��');history.back();</script>")
		response.End()
		'�ж��ڵĴ�������֤����ֵ�Ľ��������Ŀ���Ǳ��� javascript ʧЧʱ�����ݿ�����ֵ���Ӷ����³������
	end if
	SafeRequest(trim(request("shuliang")))	'SafeRequest �����Ĺ������ж�ĳ��ֵ�Ƿ�Ϊ�����ͣ����벿���� include.asp
	SafeRequest(trim(request("shichang")))
	SafeRequest(trim(request("huiyuan")))
	sql="select * from [shangpin]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs.addnew
	rs("jianjie")=trim(request("jianjie"))
	rs("riqi")=request("riqi")
	rs("mingcheng")=request("mingcheng")
	rs("shichang")=request("shichang")
	rs("huiyuan")=request("huiyuan")
	rs("xinghao")=request("xinghao")
	rs("dengji")=request("dengji")
	rs("tupian")=session("tupian")
	rs("shuoming")=request("shuoming")
	rs("beizhu")=request("beizhu")
	rs("bigclassid")=request("bigclassid")
	rs("classid")=request("classid")
	rs("shuliang")=request("shuliang")
	rs("cishu")="1"
	rs.update
	rs.close
	set rs=nothing
	session("tupian")=""	'����ʱ�洢ͼƬ·���ı�������Ϊ��
	response.Write("<script>alert('��ӳɹ���');window.location.href='addpro.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
<form name="myform" action="addpro.asp" method="post" onSubmit="return addpro();">  <tr> 
    <td align="center"><font color="#FFFFFF">����µ���Ʒ</font></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="25%">��Ʒ��飺</td>
          <td width="75%"><div align="left">
            <input name="jianjie" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>�ϼ����ڣ�</td>
          <td><div align="left">
            <input name="riqi" type="text" value="<%=now()%>" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��Ʒ���ƣ�</td>
          <td><div align="left">
            <input name="mingcheng" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>�г��۸�</td>
          <td><div align="left">
            <input name="shichang" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��Ա�۸�</td>
          <td><div align="left">
            <input name="huiyuan" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��Ʒ�ͺţ�</td>
          <td><div align="left">
            <input name="xinghao" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��ƷͼƬ��</td>
          <td><div align="left">
            <iframe src="UpFile.asp" width="260" height="20" scrolling="no" MARGINHEIGHT="0" MARGINWIDTH="0" align="bottom"></iframe>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��ϸ˵����</td>
          <td><div align="left">
            <textarea cols="60" rows="8" name="shuoming"></textarea>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��Ʒ��ע��</td>
          <td><div align="left">
            <textarea cols="60" rows="8" name="beizhu"></textarea>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td colspan="2">��Ʒ�ȼ���
            <select name="dengji">
                <option value="2">��ͨ</option>
                <option value="1">��Ʒ</option>
            </select>
<%if request("bigclassid")="" and request("classid")="" then
'�˴��жϴ���ͷ��� ID ���Ϊ�գ����ʾ��Ʒ�������δȷ������ʾ���༰��������ķ���%>
�������ࣺ
<!-- �˴�����Ĺ�����������ѡ���Ǵ���һ���¼������ҽ���Ʒ�Ĵ��� ID �ύ�� changelocation �������д��� -->
<select name="bigclassid" size="1" id="bigclassid" onChange="changelocation(document.myform.bigclassid.options[document.myform.bigclassid.selectedIndex].value)">
<!-- end -->
<%
sql="select * from [bigclass] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof	'������д���
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
����С�ࣺ
<select name="classid">
<%
sql="select * from [class] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof	'�������С��
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
<%else	'��Ϊ�գ���ʾ�Ѿ�ȷ������ͷ��࣬�����������Ʒ����𴦶�Ӧ��ֵ%>
�������ࣺ
<select name="bigclassid" size="1" id="bigclassid">
<%
sql="select * from [bigclass] where id="&request("bigclassid")&" order by paixu;"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof	'�����Ӧ�Ĵ���
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
����С�ࣺ
<select name="classid">
<%
sql="select * from [class] where id="&request("classid")&" order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3	'�����Ӧ��С��
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.close
set rs=nothing
%>
</select>
<%end if%>
&nbsp;&nbsp;��Ʒ������<input name="shuliang" type="text" size="6">
<input name="action" type="hidden" value="add">
&nbsp;&nbsp;<input type="reset" name="reset" value="��д">
&nbsp;&nbsp;<input name="submit" type="submit" value="���"></td>
          </tr>
      </table>
      <br></td>
  </tr>
</form>  
</table>
</body>
</html>