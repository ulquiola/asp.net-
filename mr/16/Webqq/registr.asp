<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�û�ע��</title>
<link href="./Css/style.css" type="text/css" rel="stylesheet">
<script src="JS/check.js"></script>
<style type="text/css">
<!--
body {
	margin-top: 0px;
	margin-bottom: 0px;
	margin-left: 0px;
	margin-right: 0px;
}
.STYLE2 {
	font-size: 9pt;
	color: #000000;
}
-->
</style>
<link href="../Css/style.css" rel="stylesheet" type="text/css">
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
.STYLE3 {
	font-size: 9pt;
	color: #666666;
}
-->
</style></head>
<script language="javascript">
var http_request = false;
function createRequest(url) 
{
	//��ʼ�����󲢷���XMLHttpRequest����
	http_request = false;
	if (window.XMLHttpRequest) { // Mozilla��������IE����������
		http_request = new XMLHttpRequest();
		if (http_request.overrideMimeType) {
			http_request.overrideMimeType("text/xml");
		}
	} else if (window.ActiveXObject) { // IE�����
		try {
			http_request = new ActiveXObject("Msxml2.XMLHTTP");
		} catch (e) {
			try {
				http_request = new ActiveXObject("Microsoft.XMLHTTP");

		   } catch (e) {}
		}
	}
	if (!http_request) {
		alert("���ܴ���XMLHTTPʵ��!");
		return false;
	}
	http_request.onreadystatechange = alertContents;    //ָ����Ӧ����
	//����HTTP����
	http_request.open("GET", url, true);
	http_request.send(null);
}
function alertContents() {    //������������ص���Ϣ
	if (http_request.readyState == 4) {//��ʾ���ݽ������
		if (http_request.status == 200) {//���ص�ǰ�����Http״̬��
			alert(http_request.responseText);
			//aaaa.innerHTML="���ݣ�"+http_request.responseText;
		} else {
			alert('�������ҳ�淢�ִ���');
		}
	}

}
</script>
<script language="javascript">
function checkName() {//�����Զ��庯��
	var UserName = myform.UserName.value;//ΪUserName������ֵ
	if(UserName=="") {//�ж��û����Ƿ�Ϊ��
		window.alert("����д�û���!");//������ʾ�Ի���
		myform.UserName.focus();//��ȡ����
		return false;
	}
	else {
		createRequest('yan.asp?UserName='+UserName);//��ת����֤ҳ��
	}
}
</script>
<body>
<table width="560" height="429"  border="0" cellpadding="-2" cellspacing="-2" background="Images/re.jpg">
  <tr>
    <td width="604" valign="top"><table width="100%" height="205"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder">
      <tr>
        <td width="10" height="32">&nbsp;</td>
        <td width="532" style="color:#FFFFFF">&nbsp;</td>
        <td width="63">&nbsp;</td>
      </tr>
      <tr valign="top">
        <td height="171" colspan="3"><table width="103%" height="129"  border="0" cellpadding="-2" cellspacing="-2">
          <tr>
            <td valign="top"><table width="103%" height="377"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder_Top">
              <tr>
                <td height="5"></td>
              </tr>
              <tr>
                <td width="349" height="263" valign="top">
				  <div align="right"></div>
				  <table width="100" border="0" align="left" cellpadding="-2" cellspacing="-2">
<tr>
<td><br><br><br><form action="register_deal.asp" method="post" name="myform">
<table width="349" height="301" border="0" align="right" cellpadding="-2" cellspacing="-2">
<tr>
<td width="21%" height="30" align="center" class="STYLE2">&nbsp;&nbsp;�û�����</td>
<td width="79%"><input name="UserName" type="text" id="UserName4" maxlength="20" class="wenbenkuang">
<a href="#" class="STYLE3" onClick="checkName()">[����û���]</a> <span class="STYLE2">*</span> </td>
</tr>
<tr>
<td height="28" align="center" class="STYLE2">&nbsp;&nbsp;��&nbsp;&nbsp;�룺</td>
<td height="28"><input name="Password1" type="password" class="wenbenkuang" id="PWD14" size="20" maxlength="20">
<span class="STYLE2">*</span></td>
</tr>
<tr>
<td height="28" align="center" class="STYLE2">&nbsp;&nbsp;ȷ�����룺</td>
<td height="28"><input name="Password2" type="password" class="wenbenkuang" id="PWD25" size="20" maxlength="20">
<span class="STYLE2"> *</span> </td>
</tr>
<tr>
<td height="28" align="center" class="STYLE2">&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;��</td>
<td class="STYLE2"><input name="sex" type="radio" class="noborder" value="��" checked>
<span class="STYLE2">��</span>&nbsp;
<input name="sex" type="radio" class="noborder" value="Ů">
Ů</td>
</tr>
<tr>
<td height="28" align="center" class="STYLE2">&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;&nbsp;�գ�</td>
<td class="word_grey"><input name="birthday" type="text" class="wenbenkuang" id="Tel">
<span class="STYLE2"> *(��ʽΪ:yyyy-mm-dd) </span></td>
</tr>
<tr>
<td height="97" align="center" class="STYLE2">&nbsp;&nbsp;ѡ��ͷ��</td>
<td class="word_grey"><table width="172" cellpadding="0" cellspacing="0">
<tr>
<td width="8" height="47"><script language="javascript">
//�鿴ȫ��ͷ��ѡ��ʱӦ�øú���
function deal(){
var someValue;
someValue=window.showModalDialog('headbrowse.asp','','dialogWidth=510px;\
dialogHeight=370px;status=no;help=no;scrollbars=no');
if (someValue == undefined){ //���û��ڵ�������ҳ�Ի�����û��ѡ��ͷ��ʱ
someValue="0";
} 
document.images.img.src="images/head/"+someValue+".gif";
myform.touxiang.value=someValue;//�������������ȡͷ��ֵ
}
</script>
</td>
<td width="97"><img src="Images/head/0.gif" name="img" width="50" height="50"></td>
<td width="245">&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td>
<td>[<a href="#" class="STYLE3" onClick="deal()">ѡ��ͷ��</a>] <input name="touxiang" type="hidden" id="toxian"></td>
<td>&nbsp;</td>
</tr>
</table></td>
</tr>
<tr>
<td height="28" align="center" class="STYLE2" style="padding-left:10px">E-mail��</td>
<td class="word_grey"><input name="Email" type="text" class="wenbenkuang" id="PWD224" size="30">
<span class="STYLE2">*</span></td>
</tr>
<tr>
<td height="34">&nbsp;</td>
<td class="word_grey"><input name="Button" type="button" class="btn_grey" value="ȷ������" onClick="mycheck()">
<input name="Submit2" type="reset" class="btn_grey" value="������д"></td>
</tr>
</table>
</form></td>
</tr>
</table> 
				</td>
                <td width="7"></td>
                <td width="204" valign="top"><p>&nbsp;</p><br>
                  <table width="149" height="173"  border="0" align="left" cellpadding="-2" cellspacing="-2">

                  <tr>
                    <td width="100%" height="173" valign="top" class="word_grey">�� �û�������ʹ��Ӣ����ĸ�����ֻ�Ӣ����ĸ�����֡��»��ߵ���ϣ����ȿ�����3-20���ַ�֮�ڡ�<br>�����գ������������գ��������������1900��3��13�������룺1900-03-13��<br>��ͷ��ͨ��ͷ�������б��ѡ��ͷ��<br>��E-mail������д��Ч��E-mail��ַ���Ա���������ϵ��</td>
                  </tr>
                  <tr align="center">
                    <td valign="middle" class="word_grey"></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td height="5"></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
