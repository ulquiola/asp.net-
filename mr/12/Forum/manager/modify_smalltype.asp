<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"-->
<%
ID=request.QueryString("ID")					'获取表单元素ID的值并赋给ID变量
if request.Form("Submit")<>"" Then				'判断表单是否提交
	ID=request.Form("ID")						'获取表单元素ID的值并赋给ID变量
	typeid=request.Form("typeid")				'获取表单元素typeid的值并赋给typeid变量
	smalltypename=request.Form("smalltypename")	'获取表单元素smalltypename的值并赋给smalltypename变量
	userid=request.Form("userid")				'获取表单元素userid的值并赋给userid变量
	account=request.Form("account")				'获取表单元素account的值并赋给account变量
	if ID<>"" Then								'判断ID值是否为空
		sql="Update tb_smalltype set typeid="&typeid&",smalltypename='"&smalltypename&"',userid="&userid&",account='"&account&"'where ID="&ID
		'应用Update语句更新指定的数据信息
		conn.execute(sql)		'执行sql语句
		response.Write("<script language='javascript'>alert('信息修改成功！');opener.location.reload();window.close();</script>")
		'弹出提示对话框
		response.End()
		'结束服务器对脚本的运行并将结果返回给浏览器
	End If
Else
	set rs=Server.CreateObject("ADODB.RecordSet")		'创建记录集
	sql="Select * From tb_user where Status='版主'"		'查询数据信息
	rs.open sql,conn,1,3								'打开记录集
	set rs_T=Server.CreateObject("ADODB.RecordSet")		'创建记录集
	sql="Select * From tb_type"							'查询数据
	rs_T.open sql,conn,1,3								'打开记录集
	set rs_1=Server.CreateObject("ADODB.RecordSet")		'创建记录集
	sql1="Select * From view_smalltype where ID="&ID	'查询数据
	rs_1.open sql1,conn,1,3								'打开记录集
End If
%>		
<html>
<head>
<title>版块信息修改</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Css/style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.STYLE1 {
	color: #000000;
	font-size: 9pt;
}
.STYLE2 {
	font-size: 10pt;
	color: #000000;
}
-->
</style></head>
<body>
<table width="496" height="205"  border="0" cellpadding="-2" cellspacing="-2">
  <tr>
    <td height="139" align="center" valign="top"><form name="form_U" method="post" action="modify_smalltype.asp">
        <table width="507" height="198" bgcolor="#FFFFFF" border="0" align="center" cellpadding="-2" cellspacing="-2" class="tabBorder">
          <tr align="center" valign="middle">
            <td height="21" background="../Images/bg_Login.GIF"><div align="center" class="STYLE2"></div></td>
            <td height="21" background="../Images/bg_Login.GIF"><div align="left"><span class="STYLE2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;=== 版块信息 === </span></div></td>
          </tr>
          <tr>
            <td width="75" height="37" align="right" valign="middle" class="STYLE1">所属讨论区：</td>
            <td width="432" valign="middle">
			<select name="typeid" id="typeid">
			<%For i=1 to rs_T.recordcount%>
            <option value="<%=rs_T("ID")%>" <%if rs_T("ID")=rs_1("typeid") then %>selected <%end if%>><%=rs_T("TypeName")%></option>
             <%
			  rs_T.movenext
			 Next
			  %>
            </select>
			</td>
          </tr>
          <tr>
            <td width="75" height="41" align="right" valign="middle" class="STYLE1">版块名称：</td>
            <td valign="middle"><input name="smalltypename" type="text" class="inputcss1" id="smalltypename" size="30" maxlength="20" value="<%=rs_1("smalltypename")%>"></td>
          </tr>
          <tr>
            <td width="75" height="30" align="right" valign="middle" class="STYLE1">版主：</td>
            <td valign="middle"><span class="word_grey">
              <select name="userid" id="userid">
			<%For i=1 to rs.recordcount%>
            <option value="<%=rs("ID")%>" <%if rs("ID")=rs_1("userid") then %>selected <%end if%>><%=rs("username")%></option>
             <%
			  rs.movenext
			 Next
			  %>
            </select>              
            </span>
			</td>
          </tr>
          <tr>
            <td width="75" height="5" align="right" valign="middle" class="STYLE1">版块说明：            </td>
            <td valign="middle"><span class="STYLE1">
             <textarea name="account" cols="40" rows="3" id="account" value="<%=rs_1("account")%>"><%=rs_1("account")%></textarea>
              <span class="word_grey">
              <input name="ID" type="hidden" id="ID" value="<%=ID%>">
            </span></span></td>
          </tr>
          <tr align="center" valign="middle">
            <td height="78" colspan="2">
                        <div align="left">
                          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                          <input name="Submit" type="submit" class="btn_grey" value="保存">
                          &nbsp;
                          <input name="Submit2" type="button" class="btn_grey" value="关闭" onClick="window.close()">
                        </div></td></tr>
        </table>
    </form></td>
  </tr>
</table>
</body>
</html>
