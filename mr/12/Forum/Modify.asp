<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"-->
<!--#include file="head.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
SQL="Select * From tb_User Where UserName='"&Session("UserName")&"'"
rs.Open SQL,Conn,1,3
%>
<html>
<head>
<title>修改用户信息！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Css/style.css" rel="stylesheet">
<script src="JS/fun.js"></script>
<style type="text/css">
<!--
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>
<body>
<table width="777" height="341"  border="0" align="center" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
  <tr>
    <td valign="top"><table width="98%"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder">
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
          <td width="13" height="32" background="Images/bg.gif">&nbsp;</td>
          <td width="689" background="Images/bg.gif" style="color:#FFFFFF"> ≡ 修改用户信息


 ≡ </td>
          <td width="73" background="Images/bg.gif">&nbsp;</td>
        </tr>
        <tr valign="top">
          <td height="171" colspan="3"><table width="100%" height="129"  border="0" cellpadding="-2" cellspacing="-2">
            <tr>
              <td valign="top">
	  			  <table width="100%" height="265"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder_Top">
			  <tr>
			  <td height="5"></td>
			  </tr>
                <tr>
                  <td width="236" height="263" valign="top"><table width="88%"  border="0" cellspacing="-2" cellpadding="-2">
                    <tr>
                      <td width="55%" height="95" align="center" class="word_grey">&nbsp;<img src="Images/reg.gif" width="84" height="54"></td>
                      <td width="45%" align="left" class="word_grey"><b>修改资料帮助</b></td>
                    </tr>
                                        <tr>
                      <td height="112" colspan="2" valign="top" class="word_grey"><ul>
                        <li>
                        用户名：可使用英文字母、数字或下划线，长度控制在3-20个字符之内。</li>
                          <li>真实姓名： 请输入真实的姓名。</li>
                          <li>密码：请设定在6-20位之间，登录密码及确认密码必须一致。</li>
                          <li>生日：输入您的生日，如果您的生日是1980年7月17日则输入：1980-07-17。 </li>
                          <li>头像：可以通过头像下拉列表框选择头像，单击“[查看全部头像]”可以浏览全部头像信息并选择。</li>
                          <li>E-mail：请填写有效的E-mail地址，以便与您联系。</li>
                      </ul>
                            </td>
                    </tr>
                    <tr align="center">
                      <td colspan="2" valign="middle" class="word_grey"></td>
                    </tr>
                  </table></td>
                  <td width="1" background="Images/line.gif"></td>
                  <td width="532" valign="top"><table width="90%"  border="0" align="center" cellpadding="-2" cellspacing="-2">
                    <tr>
                      <td><form action="Modify_U_deal.asp" method="post" name="myform">
                        <table width="100%"  border="0" cellspacing="-2" cellpadding="-2">
                          <tr>
                            <td width="18%" height="30" align="center">用 户 名：</td>
                            <td width="82%"><input name="UserName" type="text" class="inputcss1" id="UserName4" value="<%=Session("UserName")%>" maxlength="20" readonly="yes">
      *      </td>
                            </tr>
                          <tr>
                            <td height="28" align="center">真实姓名：</td>
                            <td height="28"><input name="TrueName" type="text" class="inputcss1" id="TrueName4" value="<%=rs("TrueName")%>" maxlength="10">
                              *</td>
                            </tr>
                          <tr>
                            <td height="28" align="center">密&nbsp;&nbsp;&nbsp;&nbsp;码：</td>
                            <td height="28"><input name="PWD1" type="password" class="inputcss1" id="PWD14" size="20" maxlength="20">
      *</td>
                            </tr>
                          <tr>
                            <td height="28" align="center">确认密码：</td>
                            <td height="28"><input name="PWD2" type="password" class="inputcss1" id="PWD25" size="20" maxlength="20">
                              *      </td>
                            </tr>
                          <tr>
                            <td height="28" align="center">性&nbsp;&nbsp;&nbsp;&nbsp;别：</td>
                            <td>
							<input name="sex" type="radio" class="noborder" value="男" <%If rs("Sex")="男" Then%>checked<%End If%>>
                              男&nbsp;<input name="sex" type="radio" class="noborder" value="女"<%If rs("Sex")="女" Then %>checked<%End If%>>女</td></tr>
                          <tr>
                            <td height="28" align="center">生&nbsp;&nbsp;&nbsp;&nbsp;日：</td>
                            <td class="word_grey"><input name="birthday" type="text" class="inputcss1" id="Tel" value="<%=rs("birthday")%>">
                              *（日期格式为：yyyy-mm-dd 如：1980-07-17）</td>
                          </tr>
                          <tr>
                            <td height="28" align="center">联系电话：</td>
                            <td><input name="Tel" type="text" class="inputcss1" id="Tel2" value="<%=rs("Tel")%>"></td>
                            </tr>
                          <tr>
                            <td height="28" align="center">OICQ号码：</td>
                            <td><input name="qq" type="text" class="inputcss1" id="qq2" value="<%=rs("QQ")%>"></td>
                            </tr>
                          <tr>
                            <td height="97" align="center">选择头像：</td>
                            <td class="word_grey"><table width="316" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="13" height="47">
<script language="javascript">
//通过下拉列表选择头像时应用该函数
function showlogo(){
document.images.img.src="images/head/"+document.myform.ICO.options[document.myform.ICO.selectedIndex].value;
}
//查看全部头像并选择时应用该函数
function deal(){
var someValue=window.showModalDialog('headbrowse.asp','','dialogWidth=520px;dialogHeight=430px;status=no;help=no;scrollbars=no');
document.myform.ICO.selectedIndex=someValue;
document.images.img.src="images/head/"+document.myform.ICO.options[document.myform.ICO.selectedIndex].value;
}
</script>
                                </td>
                                <td width="145"><img src="Images/head/<%=rs("ICO")%>" name="img" width="60" height="60"></td>
                                <td width="225">&nbsp;</td>
                              </tr>
                              <tr>
                                <td>&nbsp;</td>
                                <td>
								<select size="1" name="ICO" onChange="showlogo()">
                                  <%for i=0 to 16%>
                                  <option value="<%=i%>.gif" <%If rs("ICO")=i&".gif" Then%>selected<%End If%>>头像<%=i+1%>
                                  <%next%>
								</select></td>
                                <td>[<a href="#" onClick="deal()">查看全部头像</a>]</td>
                              </tr>
                            </table></td>
                          </tr>
                          <tr>
                            <td height="28" align="center" style="padding-left:10px">E-mail：</td>
                            <td class="word_grey"><input name="Email" type="text" class="inputcss1" id="PWD224" value="<%=rs("Email")%>" size="50">
                              *</td>
                          </tr>
                          <tr>
                            <td height="28" align="center">个人主页：</td>
                            <td class="word_grey"><input name="homepage" type="text" class="inputcss1" id="homepage" value="<%=rs("homepage")%>" size="50"></td>
                          </tr>
                          <tr>
                            <td height="28" align="center">家庭住址：</td>
                            <td class="word_grey"><input name="address" type="text" class="inputcss1" id="address" value="<%=rs("Address")%>" size="50"></td>
                          </tr>
                          <tr>
                            <td height="34">&nbsp;</td>
                            <td class="word_grey">
							<input name="Button" type="button" class="btn_grey" value="确定保存" onClick="check()">
                            <input name="Submit2" type="reset" class="btn_grey" value="重新填写"></td>
                          </tr>
                        </table>
					  
					  </form></td>
                    </tr>
                  </table></td>
                </tr>
				<tr>
				<td height="5"></td>
				</tr>
                </table>
			  </td>
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
<%
rs.close()
Set rs=Nothing
%>
