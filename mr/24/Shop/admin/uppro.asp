<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="../include/include.asp" -->
<!--#include file="chk_admin.asp" -->
<script language = "JavaScript">
function addpro()
{
	if(document.myform.jianjie.value=="") 
	{
		document.myform.jianjie.focus();
		alert("��������Ʒ��飡");
		return false;
	}
	if(document.myform.shuliang.value=="") 
	{
		document.myform.shuliang.focus();
		alert("��������Ʒ������");
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
}
</script>
<%
if request("action")="update" then
	if request("jianjie")="" or request("riqi")="" or request("mingcheng")="" or request("shichang")="" or request("huiyuan")="" or request("dengji")="" or request("xinghao")="" or request("shuoming")="" or request("beizhu")="" or request("bigclassid")="" or request("classid")="" or request("shuliang")="" then
		response.Write("<script>alert('����ϸ��д��');history.back();</script>")
		response.End()
	end if
	SafeRequest(trim(request("shuliang")))
	SafeRequest(trim(request("shichang")))
	SafeRequest(trim(request("huiyuan")))
	sql="select * from [shangpin] where id="&request("id")&""	'����ȡ�õ���Ʒ ID �����޸�
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs("jianjie")=trim(request("jianjie"))
	rs("riqi")=request("riqi")
	rs("shuliang")=request("shuliang")
	rs("mingcheng")=request("mingcheng")
	rs("shichang")=request("shichang")
	rs("huiyuan")=request("huiyuan")
	rs("xinghao")=request("xinghao")
	rs("dengji")=request("dengji")
	if session("tupian")<>"" then	'�˴��ж��Ƿ���Ҫ�޸�ͼƬ·��
		rs("tupian")=session("tupian")
	end if
	rs("shuoming")=request("shuoming")
	rs("beizhu")=request("beizhu")
	rs("bigclassid")=request("bigclassid")
	rs("classid")=request("classid")
	rs.update
	rs.close
	set rs=nothing
	session("tupian")=""
	response.Write("<script>alert('�޸ĳɹ���');window.location.href='lookpro.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
<form name="myform" action="uppro.asp" method="post" onSubmit="return addpro();">  <tr> 
    <td align="center"><font color="#FFFFFF">�޸���Ʒ</font></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
<%
sql="select * from [shangpin] where id="&request("id")&";"	'��ѯ��Ӧ����Ʒ��Ϣ
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
%>
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="25%">��Ʒ��飺</td>
          <td width="75%"><div align="left">
            <input name="jianjie" type="text" value="<%=rs("jianjie")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>�ϼ����ڣ�</td>
          <td><div align="left">
            <input name="riqi" type="text" value="<%=rs("riqi")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��Ʒ���ƣ�</td>
          <td><div align="left">
            <input name="mingcheng" type="text" value="<%=rs("mingcheng")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>�г��۸�</td>
          <td><div align="left">
            <input name="shichang" type="text" value="<%=rs("shichang")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��Ա�۸�</td>
          <td><div align="left">
            <input name="huiyuan" type="text" value="<%=rs("huiyuan")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��Ʒ�ͺţ�</td>
          <td><div align="left">
            <input name="xinghao" type="text" value="<%=rs("xinghao")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��ƷͼƬ��</td>
          <td><div align="left">
            <iframe src="UpFile.asp" width="300" height="20" scrolling="no" MARGINHEIGHT="0" MARGINWIDTH="0" align="bottom"></iframe>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��ϸ˵����</td>
          <td><div align="left">
            <textarea cols="50" rows="6" name="shuoming"><%=rs("shuoming")%></textarea>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��Ʒ��ע��</td>
          <td><div align="left">
            <textarea cols="50" rows="6" name="beizhu"><%=rs("beizhu")%></textarea>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td colspan="2">
��Ʒ�ȼ���
<select name="dengji">
<%if rs("dengji")="2" then	'�ж��Ƿ�Ϊ��Ʒ%>
<option value="2">��ͨ</option>
<option value="1">��Ʒ</option>
<%else%>
<option value="1">��Ʒ</option>
<option value="2">��ͨ</option>
<%end if%>
</select>
�������ࣺ
<select name="bigclassid">
<%
sql1="select * from [bigclass] where id="&rs("bigclassid")&";"	'��ѯ����Ʒ��Ӧ�Ĵ���
set rs1=Server.CreateObject("ADODB.Recordset")
rs1.open sql1,conn,3,3
%>
  <option value="<%=rs1("id")%>"><%=rs1("mingcheng")%></option>
<%
rs1.close
set rs1=nothing
%>
<%
sql2="select * from [bigclass] where id<>"&rs("bigclassid")&" order by paixu;"	'��ѯ���˶�Ӧ����Ʒ�����д���
'�����β�ѯ��Ŀ���Ǽ���������������������˲��ظ����
set rs2=Server.CreateObject("ADODB.Recordset")
rs2.open sql2,conn,3,3
do while not rs2.eof
%>
  <option value="<%=rs2("id")%>"><%=rs2("mingcheng")%></option>
<%
rs2.movenext
loop
rs2.close
set rs2=nothing
%>
</select>
����С�ࣺ
<select name="classid">
<%
sql1="select * from [class] where id="&rs("classid")&";"	'��ѯ����Ʒ��Ӧ��С��
set rs1=Server.CreateObject("ADODB.Recordset")
rs1.open sql1,conn,3,3
%>
<option value="<%=rs1("id")%>"><%=rs1("mingcheng")%></option>
<%
rs1.close
set rs1=nothing
%>
<%
sql2="select * from [class] where id<>"&rs("classid")&" order by paixu;"	'��ѯ���˶�Ӧ����Ʒ������С�࣬���ظ����
set rs2=Server.CreateObject("ADODB.Recordset")
rs2.open sql2,conn,3,3
do while not rs2.eof
%>
  <option value="<%=rs2("id")%>"><%=rs2("mingcheng")%></option>
<%
rs2.movenext
loop
rs2.close
set rs2=nothing
%>
</select>
&nbsp;&nbsp;��Ʒ������<input name="shuliang" type="text" size="6" value="<%=rs("shuliang")%>">
<input name="id" type="hidden" value="<%=request("id")%>">
<input name="action" type="hidden" value="update">
&nbsp;&nbsp;<input type="reset" name="reset" value="��д">
&nbsp;&nbsp;<input name="submit" type="submit" value="�޸�"></td>
          </tr>
<%
rs.close
set rs=nothing
%>
      </table>
      <br></td>
  </tr>
</form>  
</table>