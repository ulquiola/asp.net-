<!--#include file="Conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
<%
Set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from tb_user"
rs.open sql,conn,1,3
%>
<HTML>
<HEAD>
<title>添加好友</title>
<script language="javascript">
function mycheckadd()
{//在进行好友添加时首先需要判断选中文本域中是否有用户
if(form1.sel_place2.length<1){
	alert("请选择要添加的好友!")			//弹出提示对话框
	return 
}
var value="";								//为value赋值
for (i=0;i<form1.sel_place2.length;i++){	
	value=value+form1.sel_place2.options[i].value+",";
	//为value赋值
}
value=value.substr(0,value.length-1);
//将添加的多个用户的头像信息的值以数组的形式传递到处理页面
//其中value表示传递的是需要添加的用户ID值
//GroupID用于动态传递的ID值
window.opener.frames("mainfrm").frames("right").location = "Addfriend_deal.asp?ID=" + value + "&GroupID=<%=request.querystring("GroupID")%>"
self.close() 
}
</script>
<script language="javascript">
function allsel(n1,n2)
{//实现用户信息的移入和移出
  while(n1.selectedIndex!=-1)
  {
  	var indx=n1.selectedIndex;						//声明变量，并赋值
  	var t=n1.options[indx].text;					//声明变量，并赋值
	t1=n1.options[indx].value;						//为t1变量赋值
  	n2.options.add(new Option(t,t1));				
  	n1.remove(indx);								
  }
}
</script>
<script language="javascript"> 
function mychecktext()
//控制按钮是不可用
{ 
if(document.getElementById("content").value.length>0){ 
document.getElementById("chazhao").disabled=""; 
//设置按钮为可用状态
}else{ 
document.getElementById("chazhao").disabled="disabled"; 
//设置按钮为不可用状态
} 
} 
//查询时锁定查询结果，并用蓝色选中
function sss(myForm){
	for(i=0;i<myForm.Userlist.length;++i){
		if(myForm.Userlist.options[i].value==myForm.content.value)
		//判断选中的值和表单的内容是否相等
		{
			myForm.Userlist.options[i].selected=true;
			//将当前的内容处于选中状态
		}
		
	}
}
</script> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"><style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></HEAD>
<style type="text/css">
<!--
.STYLE3 {font-size: 9pt; color: #000000; }
.STYLE6 {color: #FF0000}
-->
</style>
<BODY style="border:none">
<form name="form1">
<table width="476" height="324" border="0" cellpadding="0" cellspacing="0" background="Images/bg.gif" id="AutoNumber1">
  <tr>
    <td colspan="3"><span class="STYLE3">&nbsp;&nbsp;&nbsp;&nbsp;<br>
      &nbsp;&nbsp;&nbsp;&nbsp;请输入用户ID或从列表中选择：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;选中的用户： </span></td>
  </tr>
  <tr>
    <td width="46%" height="281" valign="top"><br>
      <table width="221" height="26" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="24">&nbsp;</td>
        <td width="147"><input name="content" type="text"  id="content" size="18" maxlength="12" height="12 " onKeyUp="mychecktext()"></td>
        <td width="50"><input type="button" name="chazhao" value="查找" disabled="disabled" id="chazhao" onClick="sss(form1)"></td>
      </tr>
    </table>      
    &nbsp;&nbsp;  
    <select name="Userlist" size="10" multiple style="width:186px ">
            <%while not rs.EOF %>
			<%
			if Session("username1")<>rs("Username") then
			%>
            <option value="<%=rs("ID")%>"><%=rs("id")%>&nbsp;(<%=rs("Username")%>)</option>
			<%
			end if
    		rs.Movenext() 
    		wend
	%>
        </select>
          <br><br>
        </p>
        <p align="right"><br>
          <br>
&nbsp;&nbsp;&nbsp;<br>        
          <br>
      </p>      </td>
    <td width="14%" valign="top"><p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p>
        <input type="button" name="sure2" value="移入&gt;&gt;" onClick="allsel(document.form1.Userlist,document.form1.sel_place2);">
      </p>
      <p>
        <input type="button" name="sure1" value="&lt;&lt;移出" onClick="allsel(document.form1.sel_place2,document.form1.Userlist);">
      </p></td>
    <td width="40%" valign="top">
      <p><br>
          <select name="sel_place2" size="12" multiple class="wenbenkuang" id="sel_place2" style="width:178px ">
          </select>
          <br>
 </p><br><br>
      <p>         
        <input type="button" value="  添加  " name="B32" onClick="mycheckadd()">
        &nbsp;
        <input type="button" value="  取消  " name="B3" onClick="window.close()">
&nbsp;      </p></td>
  </tr>
</table>
</form>
</BODY>
</HTML>
