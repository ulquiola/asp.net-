<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理</title>
<!--#include file="../include/conn.asp" -->
<!--#include file="../include/include.asp" -->
<!--#include file="chk_admin.asp" -->

<script language = "JavaScript">
<!-- 此处代码的主要功能是为了让用户在选择商品大类的同时显示对应的分类 -->
var onecount;
onecount=0;
subcat = new Array();	//建立数组，目的是按一定规律存放所有分类
<%
i=0
sql="select * from [class] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof			''查询并且循环输出所有分类
	i=i+1	''设置数组下标，因此时ASP变量 i 的值为 1，而数组下标初始值为 0，所以下面的 i-1 就是为了符合数组下标的规则 
			''数据库中有多少符合的数据，其数组的总量就有多少，数组的最大值总是比数组的总量少 1，因为数组下标以 0 开头。 
%>
	subcat[<%=i-1%>] = new Array("<%=rs("mingcheng")%>","<%=rs("bigclassid")%>","<%=rs("id")%>");
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
onecount=<%response.Write i ''输出数组的总量，虽然ASP变量 i 在循环体外，但 i 在循环体内已经获得了最大值%>;

function changelocation(locationid)
{
	document.myform.classid.length = 0;
	var locationid=locationid;	//初始化 locationid 变量，并赋值取得的大类 ID
	var i;
	for (i=0;i < onecount; i++)		//按数组的总量（分类的总数）进行循环
	{
		if (subcat[i][1] == locationid)	//在所有分类中判断对应的大类 ID
		{ 
			 document.myform.classid.options[document.myform.classid.length] = new Option(subcat[i][0], subcat[i][2]);
			 //完成了选择商品大类的同时显示对应的分类的赋值
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
		alert("请输入商品简介！");
		return false;
	}
	if(document.myform.riqi.value=="") 
	{
		document.myform.riqi.focus();
		alert("请输入添加日期！");
		return false;
	}
	if(document.myform.mingcheng.value=="") 
	{
		document.myform.mingcheng.focus();
		alert("请输入商品名称！");
		return false;
	}
	if(document.myform.shichang.value=="") 
	{
		document.myform.shichang.focus();
		alert("请输入市场价格！");
		return false;
	}
	if(document.myform.huiyuan.value=="") 
	{
		document.myform.huiyuan.focus();
		alert("请输入会员价格！");
		return false;
	}
	if(document.myform.xinghao.value=="") 
	{
		document.myform.xinghao.focus();
		alert("请输入商品型号！");
		return false;
	}
	if(document.myform.shuoming.value=="") 
	{
		document.myform.shuoming.focus();
		alert("请输入商品说明！");
		return false;
	}
	if(document.myform.beizhu.value=="") 
	{
		document.myform.beizhu.focus();
		alert("请输入商品备注！");
		return false;
	}
	if(document.myform.shuliang.value=="") 
	{
		document.myform.shuliang.focus();
		alert("请输入商品数量！");
		return false;
	}
}
</script>
</head>
<body>
<%
if request("action")="add" then	'从隐藏表单中提取 action 值并判断是否由该表单提交的数据
	if request("jianjie")="" or request("riqi")="" or request("mingcheng")="" or request("shichang")="" or request("huiyuan")="" or request("xinghao")="" or request("dengji")="" or session("tupian")="" or request("shuoming")="" or request("beizhu")="" or request("bigclassid")="" or request("classid")="" or request("shuliang")="" then
		response.Write("<script>alert('请详细填写！');history.back();</script>")
		response.End()
		'判断内的代码是验证各个值的接收情况，目的是避免 javascript 失效时向数据库插入空值，从而导致程序错误。
	end if
	SafeRequest(trim(request("shuliang")))	'SafeRequest 函数的功能是判断某个值是否为数字型，代码部分在 include.asp
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
	session("tupian")=""	'将临时存储图片路径的变量设置为空
	response.Write("<script>alert('添加成功！');window.location.href='addpro.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
<form name="myform" action="addpro.asp" method="post" onSubmit="return addpro();">  <tr> 
    <td align="center"><font color="#FFFFFF">添加新的商品</font></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="25%">商品简介：</td>
          <td width="75%"><div align="left">
            <input name="jianjie" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>上架日期：</td>
          <td><div align="left">
            <input name="riqi" type="text" value="<%=now()%>" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>商品名称：</td>
          <td><div align="left">
            <input name="mingcheng" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>市场价格：</td>
          <td><div align="left">
            <input name="shichang" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>会员价格：</td>
          <td><div align="left">
            <input name="huiyuan" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>商品型号：</td>
          <td><div align="left">
            <input name="xinghao" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>商品图片：</td>
          <td><div align="left">
            <iframe src="UpFile.asp" width="260" height="20" scrolling="no" MARGINHEIGHT="0" MARGINWIDTH="0" align="bottom"></iframe>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>详细说明：</td>
          <td><div align="left">
            <textarea cols="60" rows="8" name="shuoming"></textarea>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>商品备注：</td>
          <td><div align="left">
            <textarea cols="60" rows="8" name="beizhu"></textarea>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td colspan="2">商品等级：
            <select name="dengji">
                <option value="2">普通</option>
                <option value="1">精品</option>
            </select>
<%if request("bigclassid")="" and request("classid")="" then
'此处判断大类和分类 ID 如果为空，则表示商品的类别尚未确定则显示大类及关联大类的分类%>
所属大类：
<!-- 此处代码的功能是在下拉选项是触发一个事件，并且将商品的大类 ID 提交到 changelocation 函数进行处理 -->
<select name="bigclassid" size="1" id="bigclassid" onChange="changelocation(document.myform.bigclassid.options[document.myform.bigclassid.selectedIndex].value)">
<!-- end -->
<%
sql="select * from [bigclass] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof	'输出所有大类
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
所属小类：
<select name="classid">
<%
sql="select * from [class] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof	'输出所有小类
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
<%else	'不为空，表示已经确定大类和分类，并且在添加商品的类别处对应赋值%>
所属大类：
<select name="bigclassid" size="1" id="bigclassid">
<%
sql="select * from [bigclass] where id="&request("bigclassid")&" order by paixu;"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof	'输出相应的大类
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
所属小类：
<select name="classid">
<%
sql="select * from [class] where id="&request("classid")&" order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3	'输出相应的小类
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.close
set rs=nothing
%>
</select>
<%end if%>
&nbsp;&nbsp;商品数量：<input name="shuliang" type="text" size="6">
<input name="action" type="hidden" value="add">
&nbsp;&nbsp;<input type="reset" name="reset" value="重写">
&nbsp;&nbsp;<input name="submit" type="submit" value="添加"></td>
          </tr>
      </table>
      <br></td>
  </tr>
</form>  
</table>
</body>
</html>