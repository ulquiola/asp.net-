<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
body,td,th {
	font-size: 12px;
}
-->
</style>
<script type="text/javascript">
function Mycheck(form){
  for(i=0;i<form.length;i++){
    if(form.elements[i].value==""){
	  alert("������������Ϣ����!");return false;}		
  }
}
</script>
<table width="755" height="365" border="0" cellspacing="0" cellpadding="0">
<form name="form_fujian" action="jmail_code.asp" method="post">
  <tr>
  	<td height="365" align="left" valign="top" background="images/bg1.jpg" id="peop"><table width="755" height="365"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="17%" align="right" valign="middle">�ռ��ˣ�</td>
        <td width="5%" align="left" valign="middle">&nbsp;</td>
        <td width="78%" align="left" valign="middle"><input name="mailto" type="text" id="mailto" size="40"></td>
      </tr>
      <tr>
        <td align="right" valign="middle">�����ˣ�</td>
        <td align="left" valign="middle">&nbsp;</td>
        <td align="left" valign="middle"><input name="mailfrom" type="text" id="mailfrom" size="40"></td>
      </tr>
      <tr>
        <td align="right" valign="middle">���룺</td>
        <td align="left" valign="middle">&nbsp;</td>
        <td align="left" valign="middle"><input name="mima" type="password" id="mima" size="44"></td>
      </tr>
      <tr>
        <td align="right" valign="middle">SMTP��������ַ��</td>
        <td align="left" valign="middle">&nbsp;</td>
        <td align="left" valign="middle"><input name="smtpaddress" type="text" id="smtpaddress" size="40"></td>
      </tr>
      <tr>
        <td align="right" valign="middle">���⣺</td>
        <td align="left" valign="middle">&nbsp;</td>
        <td align="left" valign="middle"><input name="mailsubject" type="text" id="mailsubject" size="40"></td>
      </tr>
      <tr>
        <td align="right" valign="middle">���ݣ�</td>
        <td align="left" valign="middle"><span>
        </span></td>
        <td align="left" valign="middle"><span>
          <textarea name="mailbody" cols="40" rows="5" wrap="PHYSICAL" id="mailbody"></textarea>
        </span></td>
      </tr>
      <tr>
        <td align="right" valign="middle">������</td>
        <td align="left" valign="middle">&nbsp;          </td>
        <td align="left" valign="middle">          <input name="file_path" type="text" class="input1" id="file_path" value="" size="45" maxlength="80">
          <input name="Submit3" type="button" class="input1" value="�ϴ�����" onClick="window.open('upload.asp?formname=form_fujian&editname=file_path&uppath=Upload&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=480,height=115')"></td>
      </tr>
      <tr>
        <td align="center" valign="middle">&nbsp;</td>
        <td align="center" valign="middle">&nbsp;</td>
        <td align="left" valign="middle"><input type="submit" name="Submit" value="����" onClick="return Mycheck(this.form)">
��
  <input type="reset" name="Submit2" value="��д"></td>
      </tr>
      <tr>
        <td colspan="3" align="center" valign="middle"><span>*�������ݱ�����д</span></td>
        </tr>
    </table></td>
  </tr>
  </form>
</table>

<table width="755" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="42" background="images/bottom.jpg"></td>
  </tr>
</table>
