<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file=../conn/conn.asp-->
<%
If Request.QueryString("ID")<>""then
session("ID")=Request.QueryString("ID")
end if
Set rs_personnel = Server.CreateObject("ADODB.Recordset")
sql_P="SELECT name11,birthday,sex,hy,dw,department,zw,sf,cs,phone1,phone,address,OICQ,email,family,postcode,remark FROM Tab_tongxunadd where ID="&session("ID")&""
rs_personnel.open sql_p,conn,1,3
%>
<%
if request.Form("name11")<>"" then
	name11=request.Form("name11")
	birthday=request.Form("birthday")
	sex=request.Form("sex")
	hy=request.Form("hy")
	dw=request.Form("dw")
	department=request.Form("department")
	zw=request.Form("zw")
	sf=request.Form("sf")
	cs=request.Form("cs")
	phone1=request.Form("phone1")
	phone=request.Form("phone")
	address=request.Form("address")
	OICQ=request.Form("OICQ")
	email=request.Form("email")
	family=request.Form("family")
	postcode=request.Form("postcode")
	remark=request.Form("remark")
	UP="Update Tab_tongxunadd set name11='"&name11&"',birthday='"&birthday&"',sex='"&sex&"',hy='"&hy&"',dw='"&dw&"',department='"&department&"',zw='"&zw&"',sf='"&sf&"',cs='"&cs&"',phone1='"&phone1&"',phone='"&phone&"',address='"&address&"',OICQ='"&OICQ&"',email='"&email&"',family='"&family&"',postcode='"&postcode&"',remark='"&remark&"' where ID='"&session("ID")&"'"
	conn.execute(UP)
	%>
	<script language="javascript">
	alert("信息修改成功！");
	opener.location.reload();
	window.close();
	</script>
<%end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<style type="text/css">
<!--
.style1 {
	font-size: 16px;
	font-weight: bold;
	color: #0000FF;
}
body,td,th {
	font-size: 12px;
}
body {
	background-color: #ffffff;
	margin-top: 0px;
}
.style2 {color: #FFFFFF}
.style3 {
	font-size: 14px;
	color: #000000;
}
-->
</style>
<link href="biaodan.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE4 {
	font-size: 10pt;
	color: #FFFFFF;
}
-->
</style>
<script type="text/JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
</head>
<script language="javascript">
function checkemail(email)
{
var str=email;
//在JavaScript中，正则表达式只能使用"/"开头和结束，不能使用双引号
var Expression=/\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*/;
var objExp=new RegExp(Expression);
if(objExp.test(str)==true)
{return true;}
else
{return false;}
}
</script>
<script language="javascript">
function Mycheck()
{
if(form1.name11.value=="")
{alert("请输入姓名!!");form1.name11.focus();return;}
if(form1.birthday.value=="")
{alert("出生日期不能为空!!");form1.birthday.focus();return;}
if(form1.sex.value=="")
{alert("请选择性别!!");form1.sex.focus();return;}
if(form1.hy.value=="")
{alert("请选择婚姻状况!!");form1.hy.focus();return;}
if(form1.dw.value=="")
{alert("请输入所属单位名称!!");form1.dw.focus();return;}
if(form1.department.value=="")
{alert("请输入所属部门名称!!");form1.department.focus();return;}
if(form1.zw.value=="")
{alert("请输入所属职位!!");form1.zw.focus();return;}
if(form1.sf.value=="")
{alert("请输入所属省份名称!!");form1.sf.focus();return;}
if(form1.cs.value=="")
{alert("请输入所属城市名称!!");form1.cs.focus();return;}
if(form1.phone.value=="")
{alert("请输入办公电话!!");form1.phone.focus();return;}
if(form1.phone1.value=="")
{alert("请输入移动电话!!");form1.phone1.focus();return;}
if(form1.email.value=="")
{alert("请输入E-mail!!");form1.email.focus();return;}
if(!checkemail(form1.email.value))
{alert("邮箱地址格式不正确，请重新输入!!");form1.email.focus();return;}
if(form1.postcode.value=="")
{alert("请输入邮政编码号!!");form1.postcode.focus();return;}
if(isNaN(form1.postcode.value))
{alert("邮政编码号必须为数字!!");form1.postcode.focus();return;}
if(form1.family.value=="")
{alert("请输入家庭电话!!");form1.family.focus();return;}
if(form1.address.value=="")
{alert("请输入家庭住址!!");form1.address.focus();return;}
form1.submit();
}
</script>
<body>
<form name="form1" method="post" action="update.asp">
    <script language="javascript">
	function loadCalendar(field)
	{
   var rtn = window.showModalDialog("calender.asp","","dialogWidth:280px;dialogHeight:250px;status:no;help:no;scrolling=no;scrollbars=no");
	if(rtn!=null)
		field.value=rtn;
   return;
}
</script>
  <table width="429" height="419" border="0" align="center" cellspacing="1" bgcolor="#6DBAFF">
    <tr>
      <td height="15" colspan="3"><div align="center" class="STYLE4"><span class="style1 style2 style3 style2 style2 STYLE4">修改通讯详细信息</span></div></td>
    </tr>
    <tr>
      <td width="85" bgcolor="#FFFFFF"><div align="right">姓&nbsp;&nbsp;&nbsp;&nbsp;名：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="name11" type="text" class="unnamed1" id="name11" value="<%=rs_personnel("name11")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">出生日期：</div></td>
      <td width="182" bgcolor="#FFFFFF"><input name="birthday" type="text" class="unnamed1" id="birthday" value="<%=rs_personnel("birthday")%>"></td>
      <td width="152" bgcolor="#FFFFFF"><img src="../Images/date.gif" width="20" height="20" onclick="loadCalendar(form1.birthday)"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">性&nbsp;&nbsp;&nbsp;&nbsp;别：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><select name="sex" id="sex">
	  <option value="男"<%if(instr(rs_personnel("sex"),"男")>0) then%>selected<%end if%>>男</option>
      <option value="女"<%if(instr(rs_personnel("sex"),"女")>0) then%>selected<%end if%>>女</option>
      </select>      </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">婚姻状况：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">
	  <select name="hy" id="hy">
	  <option value="已婚"<%if(instr(rs_personnel("hy"),"已婚")>0) then%>selected<%end if%>>已婚</option>
      <option value="未婚"<%if(instr(rs_personnel("hy"),"未婚")>0) then%>selected<%end if%>>未婚</option>
      </select>	  </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">所属单位：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="dw" type="text" class="unnamed1" id="dw" value="<%=rs_personnel("dw")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">所属部门：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="department" type="text" class="unnamed1" id="department" value="<%=rs_personnel("department")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">职&nbsp;&nbsp;&nbsp;&nbsp;务：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="zw" type="text" class="unnamed1" id="zw" value="<%=rs_personnel("zw")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">所在省份：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="sf" type="text" class="unnamed1" id="sf" value="<%=rs_personnel("sf")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">所在城市：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="cs" type="text" class="unnamed1" id="cs" value="<%=rs_personnel("cs")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">办公电话：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="phone1" type="text" class="unnamed1" id="phone1" value="<%=rs_personnel("phone1")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">移动电话：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="phone" type="text" class="unnamed1" id="phone" value="<%=rs_personnel("phone")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">家庭地址：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="address" type="text" class="unnamed1" id="address" value="<%=rs_personnel("address")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">OICQ：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="OICQ" type="text" class="unnamed1" id="OICQ" value="<%=rs_personnel("OICQ")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">email：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="email" type="text" class="unnamed1" id="email" value="<%=rs_personnel("email")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">家庭电话：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="family" type="text" class="unnamed1" id="family" value="<%=rs_personnel("family")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">邮政编码：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="postcode" type="text" class="unnamed1" id="postcode" value="<%=rs_personnel("postcode")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">备&nbsp;&nbsp;&nbsp;&nbsp;注：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><textarea name="remark" cols="30" rows="4" id="remark"><%=rs_personnel("remark")%>

  </textarea></td>
    </tr>
    <tr bgcolor="#F3F3F3">
      <td colspan="3"><div align="center">
        <input type="button" name="Submit" value="提交"onclick="Mycheck();">
        &nbsp;&nbsp;
        <input type="reset" name="Submit2" value="重置">&nbsp;&nbsp;
        <input name="Submit3" type="button" onClick="JScript:window.close();" value="关闭">
</div></td>
    </tr>
  </table>
  <p>&nbsp;</p>
</form>
</body>
</html>
