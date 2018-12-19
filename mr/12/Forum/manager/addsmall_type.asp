<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"--><!--包含数据库连接文件-->
<%
if request.Form("smalltypename")<>"" Then								'判断smalltypename是否为空
	typeid=request.Form("taolun")										'获取表单元素taolun的值并赋给taolun变量
	smalltypename=request.Form("smalltypename")							'获取表单元素smalltypename的值并赋给smalltypename变量
	userid=request.Form("selected")										'获取表单元素selected的值并赋给selected变量
	account=request.Form("account")										'获取表单元素account的值并赋给account变量
	if smalltypename<>"" Then											'判断smalltypename是否为空
		set rs=Server.CreateObject("ADODB.RecordSet")					'创建记录集
		sql="Select * From tb_smalltype where smalltypename='"&smalltypename&"'"	'查询数据
		rs.open sql,conn,1,3											'打开记录集
		if not(rs.eof and rs.bof) Then									'判断是否有记录
			response.Write("<script language='javascript'>alert('该版块的名称已经存在，请重新输入！');window.location.href='addsmall_type.asp';</script>")
			'弹出提示对话框
			response.End()
			'结束服务器对脚本的运行并将结果返回给浏览器
		Else
			sql="Insert Into tb_smalltype (typeid,smalltypename,userid,account) values("&typeid&",'"&smalltypename&"',"&userid&",'"&account&"')"
			'应用Insert Into实现数据信息的添加
			conn.execute(sql)
			response.Write("<script language='javascript'>alert('信息添加成功！');window.location.href='addsmall_type.asp';</script>")
			'弹出提示对话框
		End If
	End If
Else
	set rs=Server.CreateObject("ADODB.RecordSet")			'创建记录集
	sql="Select * From tb_user where Status='版主'"			'查询数据
	rs.open sql,conn,1,3									'打开记录集
	set rs_1=Server.CreateObject("ADODB.RecordSet")			'创建记录集
	sql_1="Select * From tb_Type"							'查询数据
	rs_1.open sql_1,conn,1,3								'打开记录集
End If
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Css/style.css" rel="stylesheet">
</head>
<script language="javascript">
function mycheck()//创建自定义函数
{
if(form1.taolun.value==0)														//判断是否选择了讨论区
{alert("请选择所属讨论区！");form1.taolun.focus();return false}				//弹出提示对话框
if(form1.smalltypename.value=="")												//判断版块名称是否空
{alert("请输入版块名称！");form1.smalltypename.focus();return false}			//弹出提示对话框
if(form1.selected.value==0)														//判断是否选择了版主
{alert("请选择所属版主！");form1.selected.focus();return false}				//弹出提示对话框
if(form1.account.value=="")														//判断版块说明是否为空
{alert("请输入版块的相关说明信息！");form1.account.focus();return false}		//弹出提示对话框
form1.submit();																	//提交表单
}
</script>
<body  bgcolor="B9DFFF">
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#4EAEE1">
  <form name="form1" method="post" action="addsmall_type.asp">
	    <tr>
          <td height="22" colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
            <tr>
              <td height="22" bgcolor="#4EAEE1">&nbsp;&nbsp;<font color="#FFFFFF"><strong>当前位置：添加版块</strong></font></td>
            </tr>
          </table></td>
	    </tr>
        <tr>
          <td width="20%" height="10" bgcolor="#FFFFFF"><div align="right">所属讨论区：</div></td>
          <td width="80%" bgcolor="#FFFFFF">
		  <select name="taolun" id="taolun">
              <%if rs_1.eof and rs_1.bof then%>
			  <option value="0" selected>--暂无讨论区--</option>
			     <%else%>
				 <option selected value="0">-请选择讨论区-</option>
				<%For i=1 to rs_1.recordcount%>
				 <option value="<%=rs_1("ID")%>"><%=rs_1("TypeName")%></option>
				 <%
					rs_1.movenext
					Next				
			  end if
			  %>
          </select>
		  </td>
        </tr>
        <tr>
          <td width="20%" height="11" bgcolor="#FFFFFF"><div align="right">版块名称：</div></td>
          <td bgcolor="#FFFFFF"><input name="smalltypename" type="text" class="inputcss1" id="smalltypename" size="50" maxlength="20"></td>
        </tr>
		 <tr>
          <td height="10" bgcolor="#FFFFFF"><div align="right">版主：</div></td>
          <td height="10" bgcolor="#FFFFFF"><span class="word_grey">
            <select name="selected">
              <%if rs.eof and rs.bof then%>
			  <option value="0" selected>--暂无版主--</option>
			     <%else%>
				 <option selected value="0">-请选择版主-</option>
				<%For i=1 to rs.recordcount%>
				 <option value="<%=rs("ID")%>"><%=rs("UserName")%></option>
				 <%
					rs.movenext
					Next				
			  end if
			  %>
          </select>
          </span></td>
        </tr>
		 <tr>
		   <td height="11" bgcolor="#FFFFFF"><div align="right">版块说明：</div></td>
		   <td height="11" bgcolor="#FFFFFF"><textarea name="account" cols="40" rows="3" id="account"></textarea></td>
    </tr>
		<tr>
          <td height="22" colspan="2" bgcolor="#FFFFFF"><div align="center">
            <input name="Submit3" type="button" class="btn_grey" value="保存" onClick="mycheck();">
            &nbsp;&nbsp;
            <input type="reset" class="btn_grey" value="重写">
          </div></td>
        </tr>
  </form>
</table>
</body>
</html>
