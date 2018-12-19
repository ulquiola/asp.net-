<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>请假登记--网页对话框</title>
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
if(form1.content.value=="")
{alert("请输入请假原因，再进行请假登记!!");form1.content.focus();return;}
if(form1.time1.value=="")
{alert("请输入请假时间!!");form1.time1.focus();return;}
if(form1.time2.value=="")
{alert("请输入请假返回时间!!");form1.time2.focus();return;}
form1.submit();
}
</script>
<body>
<form name="form1" action="qingjia_add_cl.asp" method="post">
<table width="627" height="436" border="0" cellpadding="0" cellspacing="0">
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
    <td valign="bottom" background="../Images/qingjia_add.gif"><table width="617" height="433" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" colspan="8">&nbsp;&nbsp;<span class="STYLE4">请假登记</span></td>
      </tr>
      <tr>
        <td height="17" colspan="8" valign="bottom"><table width="597" height="27" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="61" valign="bottom"><div align="center" class="STYLE2">姓名</div></td>
            <td width="192" valign="bottom"><label>
              <input name="name1" type="text" id="name1" onkeydown="if(event.keyCode==13){form1.department.focus();}"value="<%=session("username")%>" readonly="yes">
            </label></td>
            <td width="69" valign="bottom"><div align="center" class="STYLE2">所属部门</div></td>
            <td width="275" valign="bottom"><label>
              <input name="department" type="text" id="department" onkeydown="if(event.keyCode==13){form1.content.focus();}" />
            </label></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td width="54" height="222"><span class="STYLE2">请假原因</span></td>
        <td colspan="7"><label>
          <textarea name="content" cols="70" rows="19" wrap="physical" id="content"></textarea>
        </label></td>
      </tr>
      <tr>
        <td height="25" class="STYLE2">请假时间</td>
        <td width="65" height="25" class="STYLE2"><label></label></td>
        <td width="144" height="25" class="STYLE2"><input name="time1" type="text" id="time1" size="20" readonly="yes"></td>
        <td width="6" class="STYLE2">&nbsp;</td>
        <td width="39" height="25" class="STYLE2"><img src="../Images/date.gif" width="20" height="20" onclick="loadCalendar(form1.time1)"></td>
        <td width="20" height="25" class="STYLE2">至</td>
        <td width="145" class="STYLE2"><label>
          <input name="time2" type="text" id="time2" size="20" readonly="yes">
        </label></td>
        <td width="144" class="STYLE2">&nbsp;<img src="../Images/date.gif" width="20" height="20" onClick="loadCalendar(form1.time2)"></td>
      </tr>
      <tr>
        <td height="30" colspan="8" valign="bottom"><table width="589" height="34" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td valign="bottom"><table width="230" height="34" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td valign="middle"><div align="right"><a href="#"><img src="../Images/waichuan3.gif" width="78" height="30" border="0" onclick="Mycheck();" /></a></div></td>
                <td valign="middle"><div align="left"><a href="#"><img src="../Images/waichuan2.gif" width="81" height="29" border="0" onclick="javascript:window.close();" /></a></div></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
        </tr>
    </table></td>
  </tr>
</table>
</form>
</body>
</html>
