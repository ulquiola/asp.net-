<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../conn/conn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�����ϸ��Ϣҳ��</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
.STYLE2 {font-size: 10pt; color: #000000; }
.STYLE4 {
	font-size: 9pt;
	color: #FFFFFF;
}
-->
</style></head>
<script language="javascript">
function checkemail(email)
{
var str=email;
//��JavaScript�У�������ʽֻ��ʹ��"/"��ͷ�ͽ���������ʹ��˫����
var Expression=/\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*/;
var objExp=new RegExp(Expression);
if(objExp.test(str)==true)
{return true;}
else
{return false;}
}
</script>
<script language="javascript">
function Mycheck()
{
if(form1.name11.value=="")
{alert("����������!!");form1.name11.focus();return;}
if(form1.birthday.value=="")
{alert("�������ڲ���Ϊ��!!");form1.birthday.focus();return;}
if(form1.sex.value=="")
{alert("��ѡ���Ա�!!");form1.sex.focus();return;}
if(form1.hy.value=="")
{alert("��ѡ�����״��!!");form1.hy.focus();return;}
if(form1.dw.value=="")
{alert("������������λ����!!");form1.dw.focus();return;}
if(form1.department.value=="")
{alert("������������������!!");form1.department.focus();return;}
if(form1.zw.value=="")
{alert("����������ְλ!!");form1.zw.focus();return;}
if(form1.sf.value=="")
{alert("����������ʡ������!!");form1.sf.focus();return;}
if(form1.cs.value=="")
{alert("������������������!!");form1.cs.focus();return;}
if(form1.phone.value=="")
{alert("������칫�绰!!");form1.phone.focus();return;}
if(form1.phone1.value=="")
{alert("�������ƶ��绰!!");form1.phone1.focus();return;}
if(form1.email.value=="")
{alert("������E-mail!!");form1.email.focus();return;}
if(!checkemail(form1.email.value))
{alert("�����ַ��ʽ����ȷ������������!!");form1.email.focus();return;}
if(form1.postcode.value=="")
{alert("���������������!!");form1.postcode.focus();return;}
if(isNaN(form1.postcode.value))
{alert("��������ű���Ϊ����!!");form1.postcode.focus();return;}
if(form1.family.value=="")
{alert("�������ͥ�绰!!");form1.family.focus();return;}
if(form1.address.value=="")
{alert("�������ͥסַ!!");form1.address.focus();return;}
form1.submit();
}
</script>
<body>
<table width="814" height="505" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <script language="javascript">
	function loadCalendar(field)
	{
   var rtn = window.showModalDialog("calender.asp","","dialogWidth:280px;dialogHeight:250px;status:no;help:no;scrolling=no;scrollbars=no");
	if(rtn!=null)
		field.value=rtn;
   return;
}
</script>
    <td height="488" valign="bottom" background="../Images/main.gif"><br>
	<form name="form1" action="tongxun_xianxiadd.asp" method="post">
	  <table width="597" height="474" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td><table width="783" height="473" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="715" height="473" valign="top"><table width="732" height="90" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td height="42" colspan="2" valign="top"><table width="377" height="47" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="63" height="30" rowspan="2" align="right" valign="middle"><div align="right"><img src="../Images/add.gif" width="20" height="18" /></div></td>
                    <td width="314" height="2" valign="bottom"></td>
                  </tr>
                  <tr>
                    <td valign="middle">&nbsp;<img src="../Images/tong.gif" width="101" height="17" /></td>
                  </tr>
                </table></td>
                </tr>
              <tr>
                <td width="160" valign="bottom"><div align="right" class="STYLE2">ѡ��ͨѶ��&nbsp;</div></td>
                <td width="572" valign="bottom">
				<%
				set rs1=server.CreateObject("adodb.recordset")
				sql1="select * from Tab_tongxun"
				rs1.open sql1,conn,1,3
				session("name1")=session("name1")
				if not rs1.eof then
				max=rs1.recordcount
				%>
				<select name="name1" id="ID">
				<%
				rs1.movefirst
				while(not rs1.eof)
				%>
				<option value="<%=rs1("id")%>"><%=rs1("name1")%></option>
                <%
				rs1.movenext()
				wend
				else
				response.Write("<script language=javascript>alert('�Բ�������û�����ͨѶ��');location='javascript:history.go(-1)'</script>")
				end if
				%>
				</select>
				</td>
              </tr>
            </table>
              <table width="559" height="399" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td height="28" colspan="4"><table width="100%" height="28" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="85" height="21"><div align="center" class="STYLE2">�ա�����</div></td>
                    <td width="168" class="STYLE1"><input name="name11" type="text" id="name11" onkeydown="if(event.keyCode==13){form1.birthday.focus();}"></td>
                    <td width="62"><div align="center" class="STYLE2">��������</div></td>
                    <td width="168"><input name="birthday" type="text" id="c" readonly="yes"></td>
                    <td width="96"><div align="left"><img src="../Images/date.gif" width="20" height="20" onclick="loadCalendar(form1.birthday)"></div></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td width="84" class="STYLE2"><div align="center">�ԡ�����</div></td>
                <td width="169"><select name="sex" id="sex">
                  <option>��ѡ���Ա�</option>
				  <option value="��">��</option>
                  <option value="Ů">Ů</option>
                </select>                </td>
                <td width="62" class="STYLE2"><div align="center">����״��</div></td>
                <td width="264"><select name="hy" id="hy">
                   <option>��ѡ�����״��</option>
				  <option value="�ѻ�">�ѻ�</option>
                  <option value="δ��">δ��</option>
                </select>                </td>
              </tr>
              <tr>
                <td class="STYLE2"><div align="center">������λ</div></td>
                <td colspan="3"><input name="dw" type="text" id="dw" size="53" onkeydown="if(event.keyCode==13){form1.department.focus();}"></td>
                </tr>
              <tr>
                <td class="STYLE2"><div align="center">��������</div></td>
                <td><input name="department" type="text" id="department" onkeydown="if(event.keyCode==13){form1.zw.focus();}"></td>
                <td class="STYLE2"><div align="center">ְ������</div></td>
                <td><input name="zw" type="text" id="zw" onkeydown="if(event.keyCode==13){form1.sf.focus();}"></td>
              </tr>
              <tr>
                <td class="STYLE2"><div align="center">����ʡ��</div></td>
                <td><input name="sf" type="text" id="sf" onkeydown="if(event.keyCode==13){form1.cs.focus();}"></td>
                <td class="STYLE2"><div align="center">���ڳ���</div></td>
                <td><input name="cs" type="text" id="cs" onkeydown="if(event.keyCode==13){form1.phone.focus();}"></td>
              </tr>
              <tr>
                <td class="STYLE2"><div align="center">�칫�绰</div></td>
                <td><input name="phone" type="text" id="phone" onkeydown="if(event.keyCode==13){form1.phone1.focus();}"></td>
                <td class="STYLE2"><div align="center">�ƶ��绰</div></td>
                <td><input name="phone1" type="text" id="phone1" onkeydown="if(event.keyCode==13){form1.email.focus();}"></td>
              </tr>
              <tr>
                <td><div align="center" class="STYLE2">��������</div></td>
                <td><input name="email" type="text" id="email" onkeydown="if(event.keyCode==13){form1.postcode.focus();}"></td>
                <td><div align="center" class="STYLE2">��������</div></td>
                <td><input name="postcode" type="text" id="postcode" onkeydown="if(event.keyCode==13){form1.OICQ.focus();}"></td>
              </tr>
              <tr>
                <td height="27" class="STYLE2"><div align="center">OICQ</div></td>
                <td><input name="OICQ" type="text" id="OICQ" onkeydown="if(event.keyCode==13){form1.family.focus();}"></td>
                <td class="STYLE2"><div align="center">��ͥ�绰</div></td>
                <td><input name="family" type="text" id="family" onkeydown="if(event.keyCode==13){form1.address.focus();}"></td>
              </tr>
              <tr>
                <td height="26"><div align="center" class="STYLE2">��ͥסַ</div></td>
                <td colspan="3"><input name="address" type="text" id="address" size="53" onkeydown="if(event.keyCode==13){form1.remark.focus();}"></td>
                </tr>
              <tr>
                <td height="76"><div align="center" class="STYLE2">������ע</div></td>
                <td colspan="3"><textarea name="remark" cols="45" rows="5" id="remark"></textarea></td>
                </tr>
              <tr>
                <td height="38" colspan="4" valign="top"><table width="194" height="28" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><a href="#"><img src="../Images/10.GIF" width="80" height="31" border="0" onclick="Mycheck();"></a><a href="#"></a><a href="#"><img src="../Images/qiyebutton1.gif" width="81" height="31" border="0" onclick="JScript:form1.reset();"></a></td>
                  </tr>
                </table></td>
              </tr>
              
            </table></td>
          </tr>
        </table></td>
      </tr>
    </table>
	</form>
	</td>
  </tr>
</table>
</body>
</html>
