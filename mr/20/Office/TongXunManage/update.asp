<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file=../conn/conn.asp-->
<%
If Request.QueryString("ID")<>""then
session("ID")=Request.QueryString("ID")
end if
Set rs_personnel = Server.CreateObject("ADODB.Recordset")
sql_P="SELECT name11,birthday,sex,hy,dw,department,zw,sf,cs,phone1,phone,address,OICQ,email,family,postcode,remark FROM Tab_tongxunadd where ID="&session("ID")&""
rs_personnel.open sql_p,conn,1,3
%>
<%
if request.Form("name11")<>"" then
	name11=request.Form("name11")
	birthday=request.Form("birthday")
	sex=request.Form("sex")
	hy=request.Form("hy")
	dw=request.Form("dw")
	department=request.Form("department")
	zw=request.Form("zw")
	sf=request.Form("sf")
	cs=request.Form("cs")
	phone1=request.Form("phone1")
	phone=request.Form("phone")
	address=request.Form("address")
	OICQ=request.Form("OICQ")
	email=request.Form("email")
	family=request.Form("family")
	postcode=request.Form("postcode")
	remark=request.Form("remark")
	UP="Update Tab_tongxunadd set name11='"&name11&"',birthday='"&birthday&"',sex='"&sex&"',hy='"&hy&"',dw='"&dw&"',department='"&department&"',zw='"&zw&"',sf='"&sf&"',cs='"&cs&"',phone1='"&phone1&"',phone='"&phone&"',address='"&address&"',OICQ='"&OICQ&"',email='"&email&"',family='"&family&"',postcode='"&postcode&"',remark='"&remark&"' where ID='"&session("ID")&"'"
	conn.execute(UP)
	%>
	<script language="javascript">
	alert("��Ϣ�޸ĳɹ���");
	opener.location.reload();
	window.close();
	</script>
<%end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
<style type="text/css">
<!--
.style1 {
	font-size: 16px;
	font-weight: bold;
	color: #0000FF;
}
body,td,th {
	font-size: 12px;
}
body {
	background-color: #ffffff;
	margin-top: 0px;
}
.style2 {color: #FFFFFF}
.style3 {
	font-size: 14px;
	color: #000000;
}
-->
</style>
<link href="biaodan.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE4 {
	font-size: 10pt;
	color: #FFFFFF;
}
-->
</style>
<script type="text/JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
</head>
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
<form name="form1" method="post" action="update.asp">
    <script language="javascript">
	function loadCalendar(field)
	{
   var rtn = window.showModalDialog("calender.asp","","dialogWidth:280px;dialogHeight:250px;status:no;help:no;scrolling=no;scrollbars=no");
	if(rtn!=null)
		field.value=rtn;
   return;
}
</script>
  <table width="429" height="419" border="0" align="center" cellspacing="1" bgcolor="#6DBAFF">
    <tr>
      <td height="15" colspan="3"><div align="center" class="STYLE4"><span class="style1 style2 style3 style2 style2 STYLE4">�޸�ͨѶ��ϸ��Ϣ</span></div></td>
    </tr>
    <tr>
      <td width="85" bgcolor="#FFFFFF"><div align="right">��&nbsp;&nbsp;&nbsp;&nbsp;����</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="name11" type="text" class="unnamed1" id="name11" value="<%=rs_personnel("name11")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">�������ڣ�</div></td>
      <td width="182" bgcolor="#FFFFFF"><input name="birthday" type="text" class="unnamed1" id="birthday" value="<%=rs_personnel("birthday")%>"></td>
      <td width="152" bgcolor="#FFFFFF"><img src="../Images/date.gif" width="20" height="20" onclick="loadCalendar(form1.birthday)"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">��&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><select name="sex" id="sex">
	  <option value="��"<%if(instr(rs_personnel("sex"),"��")>0) then%>selected<%end if%>>��</option>
      <option value="Ů"<%if(instr(rs_personnel("sex"),"Ů")>0) then%>selected<%end if%>>Ů</option>
      </select>      </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">����״����</div></td>
      <td colspan="2" bgcolor="#FFFFFF">
	  <select name="hy" id="hy">
	  <option value="�ѻ�"<%if(instr(rs_personnel("hy"),"�ѻ�")>0) then%>selected<%end if%>>�ѻ�</option>
      <option value="δ��"<%if(instr(rs_personnel("hy"),"δ��")>0) then%>selected<%end if%>>δ��</option>
      </select>	  </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">������λ��</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="dw" type="text" class="unnamed1" id="dw" value="<%=rs_personnel("dw")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">�������ţ�</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="department" type="text" class="unnamed1" id="department" value="<%=rs_personnel("department")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">ְ&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="zw" type="text" class="unnamed1" id="zw" value="<%=rs_personnel("zw")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">����ʡ�ݣ�</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="sf" type="text" class="unnamed1" id="sf" value="<%=rs_personnel("sf")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">���ڳ��У�</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="cs" type="text" class="unnamed1" id="cs" value="<%=rs_personnel("cs")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">�칫�绰��</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="phone1" type="text" class="unnamed1" id="phone1" value="<%=rs_personnel("phone1")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">�ƶ��绰��</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="phone" type="text" class="unnamed1" id="phone" value="<%=rs_personnel("phone")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">��ͥ��ַ��</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="address" type="text" class="unnamed1" id="address" value="<%=rs_personnel("address")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">OICQ��</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="OICQ" type="text" class="unnamed1" id="OICQ" value="<%=rs_personnel("OICQ")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">email��</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="email" type="text" class="unnamed1" id="email" value="<%=rs_personnel("email")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">��ͥ�绰��</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="family" type="text" class="unnamed1" id="family" value="<%=rs_personnel("family")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">�������룺</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="postcode" type="text" class="unnamed1" id="postcode" value="<%=rs_personnel("postcode")%>"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="right">��&nbsp;&nbsp;&nbsp;&nbsp;ע��</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><textarea name="remark" cols="30" rows="4" id="remark"><%=rs_personnel("remark")%>

  </textarea></td>
    </tr>
    <tr bgcolor="#F3F3F3">
      <td colspan="3"><div align="center">
        <input type="button" name="Submit" value="�ύ"onclick="Mycheck();">
        &nbsp;&nbsp;
        <input type="reset" name="Submit2" value="����">&nbsp;&nbsp;
        <input name="Submit3" type="button" onClick="JScript:window.close();" value="�ر�">
</div></td>
    </tr>
  </table>
  <p>&nbsp;</p>
</form>
</body>
</html>
