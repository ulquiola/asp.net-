<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"--><!--包含数据库连接文件-->
<!--#include file="../head.asp"-->
<%
If Request.QueryString("ID")<>""then			'判断接收的ID值是否等于空
session("ID")=Request.QueryString("ID")			'为session("ID")变量赋值
end if
Set rs_personnel = Server.CreateObject("ADODB.Recordset")	'创建记录集
sql_P="SELECT UserName,TrueName,sex,birthday,Tel,qq,ICO,Email,homepage,address,Status FROM tb_User where ID="&session("ID")&""
'查询数据
rs_personnel.open sql_p,conn,1,3				'打开记录集
%>
<%
if request.Form("UserName")<>"" then			'判断接收的UserName值是否等于空
	Status=request.Form("Status")				'为Status变量赋值
	UP="Update tb_user set Status='"&Status&"' where ID="&session("ID")&""	'更新指定的记录
	conn.execute(UP)														'执行UP语句
	%>
<script langeuae="javascript">
alert("用户信息修改成功！！")					//弹出提示对话框
window.location.href="User_Manage.asp";			//跳转到指定ASP页面
</script>
<%end if%>
<%
If Request.QueryString("id") <> "" then		'判断接收的ID值是否等于空
session("id")=Request.QueryString("id")		'为session("ID")变量赋值
end if
Set rs1 = Server.CreateObject("ADODB.Recordset")	'创建记录集
sql1="SELECT UserName,TrueName,sex,birthday,Tel,qq,ICO,Email,homepage,address,Status FROM tb_User where ID="&session("ID")&""
'查询数据
rs1.open sql1,conn,1,3								'打开记录集
%>
<html>
<head>
<title>修改用户信息！</title>
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
.STYLE1 {	font-size: 10pt;
	font-weight: bold;
}
-->
</style></head>
<body bgcolor="B9DFFF">
<table width="777" height="341"  border="0" align="center" cellpadding="-2" cellspacing="-2">
  <tr>
    <td valign="top"><table width="100%" border="0" cellpadding="-2" cellspacing="-2" class="tableBorder">
      <tr>
        <td></td>
      </tr>
      <tr>
        <td></td>
      </tr>
      <tr>
        <td height="2"></td>
      </tr>
      <tr>
        <td valign="top">
      <table width="100%" height="205"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder">
        <tr>
          <td width="13" height="32" bgcolor="#4EAEE1">&nbsp;</td>
          <td width="689" bgcolor="#4EAEE1" style="color:#FFFFFF">≡<span class="STYLE1"><font color="#FFFFFF">当前位置：查看用户信息</font></span> ≡ </td>
          <td width="73" bgcolor="#4EAEE1">&nbsp;</td>
        </tr>
        <tr valign="top">
          <td height="171" colspan="3"><table width="100%" height="129"  border="0" cellpadding="-2" cellspacing="-2">
            <tr>
              <td valign="top">
	  			  <table width="100%" height="265"  border="0" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF" class="tableBorder_Top">
			  <tr>
			  <td height="5"></td>
			  </tr>
                <tr>
                 <td width="700" valign="top"><table width="90%"  border="0" align="center" cellpadding="-2" cellspacing="-2">
                    <tr>
                      <td><form action="modify_statc.asp" method="post" name="myform">
                        <table width="100%"  border="0" cellspacing="-2" cellpadding="-2">
                          <tr>
                            <td width="18%" height="30" align="center">用 户 名：</td>
                            <td width="82%"><input name="UserName" type="text" class="inputcss1" id="UserName4" value="<%=rs_personnel("UserName")%>" maxlength="20" readonly="yes">
      *      </td>
                            </tr>
                          <tr>
                            <td height="28" align="center">真实姓名：</td>
                            <td height="28"><input name="TrueName" type="text" class="inputcss1" id="TrueName4" value="<%=rs_personnel("TrueName")%>" maxlength="10" readonly="yes">
                              *</td>
                            </tr>
                          <tr>
                            <td height="28" align="center">性&nbsp;&nbsp;&nbsp;&nbsp;别：</td>
                            <td><input name="sex" type="radio" class="noborder" value="男" <%If rs_personnel("Sex")="男" Then%>checked<%End If%> readonly="yes">
                              男&nbsp;
  <input name="sex" type="radio" class="noborder" value="女"<%If rs_personnel("Sex")="女" Then %>checked<%End If%> readonly="yes">
  女</td></tr>
                          <tr>
                            <td height="28" align="center">生&nbsp;&nbsp;&nbsp;&nbsp;日：</td>
                            <td class="word_grey"><input name="birthday" type="text" class="inputcss1" id="Tel" value="<%=rs_personnel("birthday")%>" readonly="yes">
                              *（日期格式为：yyyy-mm-dd 如：1980-07-17）</td>
                          </tr>
                          <tr>
                            <td height="28" align="center">联系电话：</td>
                            <td><input name="Tel" type="text" class="inputcss1" id="Tel2" value="<%=rs_personnel("Tel")%>" readonly="yes"></td>
                            </tr>
                          <tr>
                            <td height="28" align="center">OICQ号码：</td>
                            <td><input name="qq" type="text" class="inputcss1" id="qq2" value="<%=rs_personnel("QQ")%>" readonly="yes"></td>
                            </tr>
                          <tr>
                            <td height="62" align="center">选择头像：</td>
                            <td class="word_grey"><table width="316" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="13" height="47"></td>
                                <td width="145"><img src="../Images/head/<%=rs_personnel("ICO")%>" name="img" width="60" height="60" border="1"></td>
                                <td width="225">&nbsp;</td>
                              </tr>
                            </table></td>
                          </tr>
                          <tr>
                            <td height="28" align="center" style="padding-left:10px">Email：</td>
                            <td class="word_grey"><input name="Email" type="text" class="inputcss1" id="PWD224" value="<%=rs_personnel("Email")%>" size="50" readonly="yes">
                              *</td>
                          </tr>
                          <tr>
                            <td height="28" align="center">个人主页：</td>
                            <td class="word_grey"><input name="homepage" type="text" class="inputcss1" id="homepage" value="<%=rs_personnel("homepage")%>" size="50" readonly="yes"></td>
                          </tr>
                          <tr>
                            <td height="23" align="center">家庭住址：</td>
                            <td class="word_grey"><input name="address" type="text" class="inputcss1" id="address" value="<%=rs_personnel("Address")%>" size="50" readonly="yes"			></td>
                          </tr>
                          <tr>
                            <td height="32" align="center">用户权限：</td>
                            <td class="word_grey">							
							<select name="Status" class="text" id="Status">
      <option value="普通用户" <%if(InStr(rs1("Status"),"普通用户") > 0)Then%>selected<%End IF%>>普通用户</option>
      <option value="版主" <%If(InStr(rs1("Status"),"版主") > 0)Then%>selected<%End IF%>>版主</option>
    </select>							</td>
                          </tr>
                          <tr>
                            <td height="34">&nbsp;</td>
                            <td class="word_grey"><input name="submit" type="submit" class="btn_grey" value="确定保存">
                                <input name="Submit2" type="reset" class="btn_grey" value="重新填写">
                                
                                <input name="Submit22" type="button" class="btn_grey" value="返回管理页" onClick="Jscript:window.location.href='User_Manage.asp';"></td>
                          </tr>
                        </table>
					  
					  </form></td>
                    </tr>
                  </table></td>
                </tr>
				<tr>
				<td height="5"></td>
				</tr>
              </table>			  </td>
            </tr>
          </table>
		  </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
