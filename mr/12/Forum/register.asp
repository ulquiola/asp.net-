<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<title>�û�ע�ᣡ</title>
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
          <td width="689" background="Images/bg.gif" style="color:#FFFFFF"> �� �û�ע��


 �� </td>
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
                            <td width="18%" height="30" align="center">�� �� ����</td>
                            <td width="82%"><input name="UserName" type="text" class="inputcss1" id="UserName4" maxlength="20">
      *      </td>
                            </tr>
                          <tr>
                            <td height="28" align="center">��ʵ������</td>
                            <td height="28"><input name="TrueName" type="text" class="inputcss1" id="TrueName4" maxlength="10">
                              *</td>
                            </tr>
                          <tr>
                            <td height="28" align="center">��&nbsp;&nbsp;&nbsp;&nbsp;�룺</td>
                            <td height="28"><input name="PWD1" type="password" class="inputcss1" id="PWD14" size="20" maxlength="20">
      *</td>
                            </tr>
                          <tr>
                            <td height="28" align="center">ȷ�����룺</td>
                            <td height="28"><input name="PWD2" type="password" class="inputcss1" id="PWD25" onBlur="if(this.value!=this.form.PWD1.value) alert('��������������벻һ�£�');" size="20" maxlength="20">
                              *      </td>
                            </tr>
                          <tr>
                            <td height="28" align="center">��&nbsp;&nbsp;&nbsp;&nbsp;��</td>
                            <td><input name="sex" type="radio" class="noborder" value="��">
                              ��&nbsp;
  <input name="sex" type="radio" class="noborder" value="Ů" checked>
  Ů</td></tr>
                          <tr>
                            <td height="28" align="center">��&nbsp;&nbsp;&nbsp;&nbsp;�գ�</td>
                            <td class="word_grey"><input name="birthday" type="text" class="inputcss1" id="Tel">
                              *�����ڸ�ʽΪ��yyyy-mm-dd ��</td>
                          </tr>
                          <tr>
                            <td height="28" align="center">��ϵ�绰��</td>
                            <td><input name="tel" type="text" class="inputcss1" id="tel"></td>
                          </tr>
                          <tr>
                            <td height="28" align="center">OICQ���룺</td>
                            <td><input name="qq" type="text" class="inputcss1" id="qq"></td>
                            </tr>
                          <tr>
                            <td height="97" align="center">ѡ��ͷ��</td>
                            <td class="word_grey"><table width="253" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="13" height="47">
<script language="javascript">
//ͨ�������б�ѡ��ͷ��ʱӦ�øú���
function showlogo(){
document.images.img.src="images/head/"+
document.myform.ICO.options[document.myform.ICO.selectedIndex].value;
}
//�鿴ȫ��ͷ��ѡ��ʱӦ�øú���
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
                                  <option value="<%=i-1%>.gif">ͷ��<%=i%>
                                  <%next%>
                                  </select></td>
                                <td>[<a href="#" onClick="deal()" style="text-decoration:none">�鿴ȫ��ͷ��</a>]</td>
                              </tr>
                            </table></td>
                          </tr>
                          <tr>
                            <td height="28" align="center" style="padding-left:10px">E-mail��</td>
                            <td class="word_grey"><input name="Email" type="text" class="inputcss1" id="PWD224" size="50">
                              *</td>
                          </tr>
                          <tr>
                            <td height="28" align="center">������ҳ��</td>
                            <td class="word_grey"><input name="homepage" type="text" class="inputcss1" id="homepage" size="50"></td>
                          </tr>
                          <tr>
                            <td height="28" align="center">��ͥסַ��</td>
                            <td class="word_grey"><input name="address" type="text" class="inputcss1" id="address" size="50"></td>
                          </tr>
                          <tr>
                            <td height="34">&nbsp;</td>
                            <td class="word_grey"><input name="Button" type="button" class="btn_grey" value="ȷ������" onClick="check();">
                                <input name="Submit2" type="reset" class="btn_grey" value="������д"></td>
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
                    <td width="55%" align="left" class="word_grey"><b>ע�����</b></td>
                    </tr>
                    <tr>
                      <td height="112" colspan="2" align="left" valign="top" class="word_grey"><ul>
                        <li>
                          <div align="left">
                            �û�������ʹ��Ӣ����ĸ�����ֻ��»��ߣ����ȿ�����3-20���ַ�֮�ڡ�</div>
                        </li>
                          <li>
                            <div align="left">��ʵ������ ��������ʵ��������</div>
                          </li>
                          <li>
                            <div align="left">���룺���趨��6-20λ֮�䣬��¼���뼰ȷ���������һ�¡�</div>
                          </li>
                          <li>
                            <div align="left">���գ������������գ��������������1980��7��17�������룺1980-07-17�� </div>
                          </li>
                          <li>
                            <div align="left">ͷ�񣺿���ͨ��ͷ�������б��ѡ��ͷ�񣬵�����[�鿴ȫ��ͷ��]���������ȫ��ͷ����Ϣ��ѡ��</div>
                          </li>
                          <li>
                            <div align="left">E-mail������д��Ч��E-mail��ַ���Ա�������ϵ��</div>
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
