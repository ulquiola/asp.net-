<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.style1 {color: #f2ab5b}
.style3 {font-weight: bold}
.style5 {
	color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
<!--#include file="include/conn.asp" -->
<!--#include file="include/include.asp" -->
<!--#include file="include/user_include.asp" -->
<table width="792" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="3"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="792" height="165" background="images/index_r1_c1.jpg"><!--#include file="top.asp" --></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td width="197"><!--#include file="left.asp" --></td>
    <td width="590" valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="590">
  <tr>
    <td colspan="3"><img name="index_7_r1_c1" src="images/yhzx.jpg" width="590" height="49" border="0" alt=""></td>
  </tr>
  <tr>
    <td colspan="3"><img name="index_7_r2_c1" src="images/index_7_r2_c1.jpg" width="590" height="9" border="0" alt=""></td>
  </tr>
  <tr>
    <td width="16" height="643" background="images/index_7_r3_c1.jpg">&nbsp;</td>
    <td width="565" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>&nbsp;</td>
      </tr>
    </table>
      <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><div align="center"><a href="usercenter.asp?action=1">�û���Ϣ</a></div></td>
          <td><div align="center"><a href="usercenter.asp?action=2">�޸���Ϣ</a></div></td>
          <td><div align="center"><a href="usercenter.asp?action=3">�޸�����</a></div></td>
          <td><div align="center"><a href="usercenter.asp?action=4">�����һ�</a></div></td>
          <td><div align="center"><a href="usercenter.asp?action=5">�û�����</a></div></td>
        </tr>
      </table>
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table>
<%
action=request("action")
if request("action")="" then	'���û�û�н��в���ʱ
	action="1"
end if
select case action
	case "1"
		lookuser()	'�鿴�û���Ϣ
	case "2"
		upuser()	'�޸��û���Ϣ������
	case "3"
		uppass()	'�޸��û����������
	case "4"
		lookpass()	'�����һ��ύ��
	case "5"
		dingdan()	'��ʾ�û����ж���
	case "updateuser"
		updateuser()	'�޸��û���Ϣ
	case "updatepass"
		updatepass()	'�޸��û�����
	case "thispass"
		thispass()	'�û������һ�
end select
%>
	  </td>
    <td width="10" background="images/index_7_r3_c3.jpg">&nbsp;</td>
  </tr>
  <tr>
    <td height="7" colspan="3"><img name="index_7_r4_c1" src="images/index_7_r4_c1.jpg" width="590" height="7" border="0" alt=""></td>
  </tr>
</table></td>
    <td width="8"><img name="index_r2_c3" src="images/index_r2_c3.jpg" width="5" height="753" border="0" alt=""></td>
  </tr>
  <tr>
    <td colspan="3"><!--#include file="foot.asp" --></td>
  </tr>
</table>