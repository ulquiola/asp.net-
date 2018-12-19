<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>上下班登记--网页对话框</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
.STYLE2 {
	font-size: 9pt;
	color: #000000;
}
.STYLE4 {
	font-size: 10pt;
	color: #FFFFFF;
}
-->
</style></head>
<script language="javascript">
function Mycheck()
{
if(form1.name1.value=="")
{alert("请输入姓名!!");form1.name1.focus();return;}
if(form1.department.value=="")
{alert("请输入所属部门!!");form1.department.focus();return;}
if(form1.enroltype.value=="")
{alert("请选择登记类型!!");form1.enroltype.focus();return;}
if(form1.enrolremark.value=="")
{alert("请输入登记备注!");form1.enrolremark.focus();return;}
form1.submit();
}
</script>
<body>
<form name="form1" action="onduty_add_cl.asp" method="post">
<table width="627" height="329" border="0" cellpadding="0" cellspacing="0">
  <script language="javascript">
	function loadCalendar(field)
	{
   var rtn = window.showModalDialog("calender.asp","","dialogWidth:280px;dialogHeight:250px;status:no;help:no;scrolling=no;scrollbars=no");
	if(rtn!=null)
		field.value=rtn;
   return;
}
</script>
  <tr>
    <td valign="top" background="../Images/ondutybei.gif"><table width="617" height="312" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" colspan="8" valign="bottom">&nbsp;&nbsp;<span class="STYLE4">上下班登记</span></td>
      </tr>
      <tr>
        <td height="17" colspan="8" valign="bottom"><table width="597" height="39" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="61" valign="bottom"><div align="center" class="STYLE2">姓名</div></td>
            <td width="192" valign="bottom"><label>
              <input name="name1" type="text" id="name1" onkeydown="if(event.keyCode==13){form1.department.focus();}"value="<%=session("username")%>" readonly="yes">
            </label></td>
            <td width="69" valign="bottom"><div align="center" class="STYLE2">所属部门</div></td>
            <td width="275" valign="bottom"><label>
              <input name="department" type="text" id="department" onkeydown="if(event.keyCode==13){form1.enroltype.focus();}">
            </label></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td width="60" height="222" valign="top"><table width="61" height="105" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td height="38" class="STYLE2">登记类型</td>
          </tr>
          <tr>
            <td height="59" class="STYLE2">登记备注</td>
          </tr>
        </table></td>
        <td colspan="7" valign="top"><table width="548" height="107" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td height="38"><select name="enroltype" id="enroltype" onkeydown="if(event.keyCode==13){form1.enrolremark.focus();}">
              <option>请选择登记类型类</option>
              <option value="上班登记">上班登记</option>
              <option value="下班登记">下班登记</option>
            </select>
            </td>
          </tr>
          <tr>
            <td height="59"><textarea name="enrolremark" cols="40" rows="6" id="enrolremark"></textarea></td>
          </tr>
        </table>        <label></label>
        <table width="404" height="60" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td valign="bottom"><table width="230" height="34" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right"><a href="#"><img src="../Images/waichuan6.gif" width="77" height="29" border="0" onclick="Mycheck();" /></a></div></td>
                  <td valign="middle"><div align="left"><a href="#"><img src="../Images/waichuan2.gif" width="81" height="29" border="0" onclick="javascript:window.close();" /></a></div></td>
                </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="25" colspan="8" class="STYLE2"><label></label>          <label></label>&nbsp;</td>
        </tr>
      <tr>
        <td height="13" colspan="8" valign="bottom">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</form>
</body>
</html>
