<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"--><!--包含数据库连接文件-->
<%
if request.Form("TypeName")<>"" Then	'获取TypeName值
	TName=request.Form("TypeName")		'获取表单元素TypeName的值并赋给TName变量
	owner=request.Form("selected")		'获取表单元素selected的值并赋给owner变量
	if TName<>"" Then					'判断TName是否有值
		set rs=Server.CreateObject("ADODB.RecordSet")				'创建记录集
		sql="Select * From tb_Type where TypeName='"&TName&"'"		'查询数据
		rs.open sql,conn,1,3										'打开记录集
		if not(rs.eof and rs.bof) Then								'判断是否有记录
			response.Write("<script language='javascript'>alert('该讨论区的名称已经存在，请重新输入！');window.location.href='Add_Type.asp';</script>")
			'弹出提示对话框
			response.End()
			'结束服务器对脚本的运行并将结果返回给浏览器
		Else
			sql="Insert Into tb_Type (TypeName,owner) values('"&TName&"',"&owner&")"
			'应用Insert Into语句添加数据信息
			conn.execute(sql)
			'执行sql语句
			response.Write("<script language='javascript'>alert('信息添加成功！');window.location.href='Add_Type.asp';</script>")
			'弹出提示对话框，并跳转到指定的ASP页面
		End If
	Else
		response.Write("<script language='javascript'>alert('请输入讨论区的名称！')</script>")
		'弹出提示对话框
	End If
Else
	set rs=Server.CreateObject("ADODB.RecordSet")			'创建记录集
	sql="Select * From tb_user where Status='版主'"			'查询数据
	rs.open sql,conn,1,3									'打开记录集
End If
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Css/style.css" rel="stylesheet">
</head>
<script language="javascript">
function mycheck()
{//创建自定义函数
if(form_U.TypeName.value=="")
{alert("讨论区的名称不能为空！");form_U.TypeName.focus();return false}
if(form_U.selected.value==0)
{alert("请选择讨论区的所属的版主！");form_U.selected.focus();return false}
form_U.submit();
}
</script>
<body  bgcolor="B9DFFF">
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#4EAEE1">
  <form name="form_U" method="post" action="Add_Type.asp">
	    <tr>
          <td height="22" colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
            <tr>
              <td height="22" bgcolor="#4EAEE1">&nbsp;&nbsp;<font color="#FFFFFF"><strong>当前位置：添加讨论区</strong></font></td>
            </tr>
          </table></td>
	    </tr>
        <tr>
          <td width="20%" height="22" bgcolor="#FFFFFF"><div align="right">讨论区名称：</div></td>
          <td width="80%" bgcolor="#FFFFFF">&nbsp;
          <input name="TypeName" type="text" class="inputcss1" id="TypeName" maxlength="20"></td>
        </tr>
		 <tr>
          <td height="22" bgcolor="#FFFFFF"><div align="right">版主：</div></td>
          <td height="22" bgcolor="#FFFFFF">&nbsp;<span class="word_grey">
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
