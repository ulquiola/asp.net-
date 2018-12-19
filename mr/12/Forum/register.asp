<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<title>用户注册！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Css/style.css" rel="stylesheet">
<script src="JS/fun.js"></script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>
<body>
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
      </tr>
      <tr>
        <td valign="top">
      <table width="100%" height="205"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder">
        <tr>
          <td width="13" height="32" background="Images/bg.gif">&nbsp;</td>
          <td width="689" background="Images/bg.gif" style="color:#FFFFFF"> ≡ 用户注册


 ≡ </td>
          <td width="73" background="Images/bg.gif">&nbsp;</td>
        </tr>
        <tr valign="top">
          <td height="171" colspan="3"><table width="100%" height="129"  border="0" cellpadding="-2" cellspacing="-2">
            <tr>
              <td valign="top">
	  			  <table width="765" height="265"  border="0" align="right" cellpadding="-2" cellspacing="-2" class="tableBorder_Top">
			  <tr>
			  <td height="5"></td>
			  </tr>
                <tr>
                  <td width="508" height="263" valign="top">
				  <table width="472" border="0" align="center" cellpadding="-2" cellspacing="-2">
                    <tr>
                      <td><form action="register_deal.asp" method="post" name="myform">
                        <table width="100%"  border="0" cellspacing="-2" cellpadding="-2">
                          <tr>
                            <td width="18%" height="30" align="center">用 户 名：</td>
                            <td width="82%"><input name="UserName" type="text" class="inputcss1" id="UserName4" maxlength="20">
      *      </td>
                            </tr>
                          <tr>
                            <td height="28" align="center">真实姓名：</td>
                            <td height="28"><input name="TrueName" type="text" class="inputcss1" id="TrueName4" maxlength="10">
                              *</td>
                            </tr>
                          <tr>
                            <td height="28" align="center">密&nbsp;&nbsp;&nbsp;&nbsp;码：</td>
                            <td height="28"><input name="PWD1" type="password" class="inputcss1" id="PWD14" size="20" maxlength="20">
      *</td>
                            </tr>
                          <tr>
                            <td height="28" align="center">确认密码：</td>
                            <td height="28"><input name="PWD2" type="password" class="inputcss1" id="PWD25" onBlur="if(this.value!=this.form.PWD1.value) alert('您两次输入的密码不一致！');" size="20" maxlength="20">
                              *      </td>
                            </tr>
                          <tr>
                            <td height="28" align="center">性&nbsp;&nbsp;&nbsp;&nbsp;别：</td>
                            <td><input name="sex" type="radio" class="noborder" value="男">
                              男&nbsp;
  <input name="sex" type="radio" class="noborder" value="女" checked>
  女</td></tr>
                          <tr>
                            <td height="28" align="center">生&nbsp;&nbsp;&nbsp;&nbsp;日：</td>
                            <td class="word_grey"><input name="birthday" type="text" class="inputcss1" id="Tel">
                              *（日期格式为：yyyy-mm-dd ）</td>
                          </tr>
                          <tr>
                            <td height="28" align="center">联系电话：</td>
                            <td><input name="tel" type="text" class="inputcss1" id="tel"></td>
                          </tr>
                          <tr>
                            <td height="28" align="center">OICQ号码：</td>
                            <td><input name="qq" type="text" class="inputcss1" id="qq"></td>
                            </tr>
                          <tr>
                            <td height="97" align="center">选择头像：</td>
                            <td class="word_grey"><table width="253" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="13" height="47">
<script language="javascript">
//通过下拉列表选择头像时应用该函数
function showlogo(){
document.images.img.src="images/head/"+
document.myform.ICO.options[document.myform.ICO.selectedIndex].value;
}
//查看全部头像并选择时应用该函数
function deal(){
var someValue=window.showModalDialog('headbrowse.asp','','dialogWidth=520px;\
dialogHeight=430px;status=no;help=no;scrollbars=no');
document.myform.ICO.selectedIndex=someValue;
document.images.img.src="images/head/"+
document.myform.ICO.options[document.myform.ICO.selectedIndex].value;
}
</script>                                </td>
                                <td width="145"><img src="Images/head/0.gif" name="img" width="60" height="60"></td>
                                <td width="225">&nbsp;</td>
                              </tr>
                              <tr>
                                <td>&nbsp;</td>
                                <td><select size="1" name="ICO" onChange="showlogo()">
                                  <%for i=1 to 17%>
                                  <option value="<%=i-1%>.gif">头像<%=i%>
                                  <%next%>
                                  </select></td>
                                <td>[<a href="#" onClick="deal()" style="text-decoration:none">查看全部头像</a>]</td>
                              </tr>
                            </table></td>
                          </tr>
                          <tr>
                            <td height="28" align="center" style="padding-left:10px">E-mail：</td>
                            <td class="word_grey"><input name="Email" type="text" class="inputcss1" id="PWD224" size="50">
                              *</td>
                          </tr>
                          <tr>
                            <td height="28" align="center">个人主页：</td>
                            <td class="word_grey"><input name="homepage" type="text" class="inputcss1" id="homepage" size="50"></td>
                          </tr>
                          <tr>
                            <td height="28" align="center">家庭住址：</td>
                            <td class="word_grey"><input name="address" type="text" class="inputcss1" id="address" size="50"></td>
                          </tr>
                          <tr>
                            <td height="34">&nbsp;</td>
                            <td class="word_grey"><input name="Button" type="button" class="btn_grey" value="确定保存" onClick="check();">
                                <input name="Submit2" type="reset" class="btn_grey" value="重新填写"></td>
                          </tr>
                        </table>
					  
					  </form></td>
                    </tr>
                  </table>
				  </td>
                  <td width="257" align="right" valign="top">
				  
				  <div style="width:80%; height:500; border:dotted #000000 1px; background-color:#E9EBEB">
				  <table height="188"  border="0" cellpadding="-2" cellspacing="-2">
                    <tr>
                    <td width="45%" height="82" align="center" class="word_grey">&nbsp;<img src="Images/help.gif" width="55" height="55"></td>
                    <td width="55%" align="left" class="word_grey"><b>注册帮助</b></td>
                    </tr>
                    <tr>
                      <td height="112" colspan="2" align="left" valign="top" class="word_grey"><ul>
                        <li>
                          <div align="left">
                            用户名：可使用英文字母、数字或下划线，长度控制在3-20个字符之内。</div>
                        </li>
                          <li>
                            <div align="left">真实姓名： 请输入真实的姓名。</div>
                          </li>
                          <li>
                            <div align="left">密码：请设定在6-20位之间，登录密码及确认密码必须一致。</div>
                          </li>
                          <li>
                            <div align="left">生日：输入您的生日，如果您的生日是1980年7月17日则输入：1980-07-17。 </div>
                          </li>
                          <li>
                            <div align="left">头像：可以通过头像下拉列表框选择头像，单击“[查看全部头像]”可以浏览全部头像信息并选择。</div>
                          </li>
                          <li>
                            <div align="left">E-mail：请填写有效的E-mail地址，以便与您联系。</div>
                          </li>
                      </ul>                            </td>
                    </tr>
                  </table>
				  </div>
				  
				  </td>
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
