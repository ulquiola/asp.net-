<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"--><!--包含数据库连接文件-->
<%
TypeID=request.QueryString("TypeID")								'获取TypeID的值
if request.Form("Submit")<>"" Then									'判断表单是否提交
	TypeID=request.Form("TypeID")									'获取表单元素TypeID的值并赋给TypeID变量
	owner=request.Form("owner")										'获取表单元素owner的值并赋给owner变量
	if TypeID<>"" Then												'判断TypeID是否为空
		sql="Update tb_Type set owner="&owner&" where ID="&TypeID	'更新指定记录
		conn.execute(sql)											'执行sql语句
		response.Write("<script language='javascript'>alert('信息修改成功！');opener.location.reload();window.close();</script>")
		'弹出提示信息对话框
		response.End()
		'结束服务器对脚本的运行并将结果返回给浏览器
	End If
Else
	set rs=Server.CreateObject("ADODB.RecordSet")					'创建记录集
	sql="Select * From tb_user where Status='版主'"					'查询数据
	rs.open sql,conn,1,3											'打开记录集
	set rs_T=Server.CreateObject("ADODB.RecordSet")					'创建记录集
	sql="Select * From tb_Type where ID="&TypeID					'查询数据
	rs_T.open sql,conn,1,3											'打开记录集
End If
%>
<html>
<head>
<title>讨论区信息修改</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="./Css/style.css" rel="stylesheet">
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
-->
</style></head>
<body>
<table width="240" height="139"  border="0" align="center" cellpadding="-2" cellspacing="-2">
  <tr>
    <td height="139" align="center"><form name="form_U" method="post" action="modify_Type.asp">
        <table width="259" height="124" bgcolor="#FFFFFF" border="0" align="center" cellpadding="-2" cellspacing="-2" class="tabBorder">
          <tr align="center" valign="middle">
            <td height="25" colspan="2" background="../Images/bg_Login.GIF"><span class="STYLE1">=== 讨论区信息 === </span> </td>
          </tr>
          <tr>
            <td width="78" height="40" align="right" valign="middle" class="STYLE1">讨论区名称：</td>
            <td width="168" valign="middle"><input name="TypeName" type="text" id="TypeName" value="<%=rs_T("TypeName")%>" maxlength="20" readonly="yes"></td>
          </tr>
          <tr>
            <td height="30" align="right" valign="middle" class="STYLE1">版主名称：</td>
            <td width="168" valign="middle"><span class="word_grey">
              <select name="owner">
                <%For i=1 to rs.recordcount%>
                <option value="<%=rs("ID")%>" <%if rs("ID")=rs_T("owner") then %>selected <%end if%>><%=rs("UserName")%></option>
                <%
			  rs.movenext
			  Next
			  %>
              </select>
              <input name="TypeID" type="hidden" id="TypeID" value="<%=TypeID%>">
</span></td>
          </tr>
          <tr align="center" valign="middle">
            <td height="27" colspan="2"><input name="Submit" type="submit" class="btn_grey" value="保存">
&nbsp;
            <input name="Submit2" type="button" class="btn_grey" value="关闭" onClick="window.close()"></td>
          </tr>
        </table>
    </form></td>
  </tr>
</table>
</body>
</html>
