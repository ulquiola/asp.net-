<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"-->
<%
	typeid=request.querystring("TypeID")				'��ȡ��Ԫ��ID��ֵ������ID����
	set rs_T=Server.CreateObject("ADODB.RecordSet")		'������¼��
	sql="Select * From tb_type"							'��ѯ����
	rs_T.open sql,conn,1,3								'�򿪼�¼��
	set rs_1=Server.CreateObject("ADODB.RecordSet")		'������¼��
	sql1="Select * From view_smalltype where ID="&typeid'��ѯ����
	rs_1.open sql1,conn,1,3								'�򿪼�¼��

Set rs_T=Server.CreateObject("ADODB.RecordSet")			'������¼��
SQL="Select * From tb_smalltype"						'��ѯ����
rs_T.Open SQL,Conn,1,3									'�򿪼�¼��
Set rs=Server.CreateObject("ADODB.RecordSet")			'������¼��
SQL="Select * From tb_User Where UserName='"&Session("UserName")&"'"	'��ѯ����
rs.Open SQL,Conn,1,3													'�򿪼�¼��
If rs.Eof and rs.Bof Then												'�ж��Ƿ����û���Ϣ
%>
<script language="javascript">
	alert("��ע���Ϊ�û����ſ��Է���������Ϣ��");						//������ʾ�Ի���
	window.location.href="index.asp";									//��ת��ָ����ҳ��
</script>
<%
 Response.End()															'�������
 End If
 %>
<html>
<head>
<title>�������⣡</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Css/style.css" rel="stylesheet">
<script src="JS/UBBCode.asp"></script>
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
<script language="javascript">
function check()
{
	if(form1.subject.value==""){
		alert("������Ϣ������Ϊ�գ�");form1.subject.focus();return;
	}
	if(form1.content.value==""){
		alert("�������ݲ�����Ϊ�գ�");form1.content.focus();return;
	}
	form1.submit();
}
</script>
<SCRIPT language=JavaScript>
<!--
var LastCount =0;
function CountStrByte(Message,Total,Used,Remain){ //�ֽ�ͳ��
 var ByteCount = 0;
 var StrValue  = Message.value;
 var StrLength = Message.value.length;
 var MaxValue  = Total.value;

 if(LastCount != StrLength) { // �ڴ��жϣ�����ѭ������
	for (i=0;i<StrLength;i++){
		ByteCount  = (StrValue.charCodeAt(i)<=256) ? ByteCount + 1 : ByteCount + 2;
      if (ByteCount>MaxValue) {
      	Message.value = StrValue.substring(0,i);
			alert("����������಻�ܳ��� " +MaxValue+ " ���ֽڣ�\nע�⣺һ������Ϊ���ֽڡ�");
         ByteCount = MaxValue;
         break;
      }
	}
   Used.value = ByteCount;
   Remain.value = MaxValue - ByteCount;
   LastCount = StrLength;
 }
}
//-->
</SCRIPT>
<body>
<table width="777" height="341"  border="0" align="center" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
  <tr>
    <td valign="top"><table width="100%" height="562" border="0" align="center" cellpadding="-2" cellspacing="-2" class="tableBorder">
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
          <td width="689" background="Images/bg.gif" style="color:#FFFFFF"> �� ��������


 �� </td>
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
                  <td width="270" height="263" valign="top"><div align="center">
                    <p>
                       === ��������Ϣ ===<br>
                      <%=Session("UserName")%></p>
					  <%
					  UserName=rs("UserName")
					  Email=rs("Email")
					  homepage=rs("homepage")
					  ICO=rs("ICO")
					  QQ=rs("QQ")
					  %>
                    <p><img src="Images/head/<%=ICO%>" width="60" height="60" border=0> </p>
                    <p>���ǣ�<%=rs("sex")%>��</p>
                    <p>
                        <img src="Images/email.gif" alt="<%=UserName%>��Email�ǣ�" width="45" height="16"> <%=Email%></p>
                    <p>
					    <img src="Images/oicq.gif" alt="<%=UserName%>��QQ�����ǣ�" width="16" height="16"> QQ:<%=QQ%></p>
                    <p>
					      <img src="Images/ip.gif" alt="IPΪ��" width="16" height="15"> IP��<%=Request.ServerVariables("REMOTE_ADDR")%></p>
                  </div></td>
                  <td width="23" background="Images/line.gif"></td>
                  <td width="661" valign="top">
				  <form action="add_Topic_deal.asp" method="post" name="form1">
				  <table width="100%"  border="0" cellspacing="-2" cellpadding="-2">
                    <tr>
                      <td height="34" align="center" style="padding-left:10px">���</td>
                      <td class="word_grey">
<select name="TypeName">
<%For i=1 to rs_T.recordcount%>
<option value="<%=rs_T("ID")%>" <%if trim(rs_T("ID"))=typeid then %>selected <%end if%>><%=rs_T("smalltypename")%></option>
					  <%
					  rs_T.movenext
					  Next
					  %>
                      </select>
					  </td>
                    </tr>
                    <tr>
                      <td height="34" align="center" style="padding-left:10px">���⣺</td>
                      <td class="word_grey"><input name="subject" type="text" id="subject" size="50"></td>
                    </tr>
                    <tr>
                      <td width="16%" height="90" align="center">���飺</td>
                      <td width="84%"><INPUT name=emote type=radio class="noborder" value=0 checked>
                        <IMG height=20 
            src="Images/Face/face0.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=1>
                        <IMG height=19 
            src="Images/Face/face1.gif" width=19 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=2>
                        <IMG height=20 
            src="Images/Face/face2.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=3>
                        <IMG height=20 
            src="Images/Face/face3.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=4>
                        <IMG height=20 
            src="Images/Face/face4.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=5>
                        <IMG height=20 
            src="Images/Face/face5.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=6>
                        <IMG height=20 
            src="Images/Face/face6.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=7>
                        <img height=20 
            src="Images/Face/face7.gif" width=20 align=middle border=0><br>
                        <INPUT name=emote type=radio class="noborder" value=8>
                        <IMG height=20 
            src="Images/Face/face8.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=9>
                        <IMG height=20 
            src="Images/Face/face9.gif" width=20 align=middle border=0> 
			            <INPUT name=emote type=radio class="noborder" value=10>
                        <IMG height=20 
            src="Images/Face/face10.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=11>
                        <IMG height=20 
            src="Images/Face/face11.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=12>
                        <IMG height=20 
            src="Images/Face/face12.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=13>
                        <IMG height=20 
            src="Images/Face/face13.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=14>
                        <IMG height=20 
            src="Images/Face/face14.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=15>
                        <img height=20 
            src="Images/Face/face15.gif" width=20 align=middle border=0><br>
                        <INPUT name=emote type=radio class="noborder" value=16>
                        <IMG height=20 
            src="Images/Face/face16.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=17>
                        <IMG height=20 
            src="Images/Face/face17.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=18>
                        <IMG height=20 
            src="Images/Face/face18.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=19>
                        <IMG height=19 
            src="Images/Face/face19.gif" width=19 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=20>
                        <IMG height=19 
            src="Images/Face/face20.gif" width=19 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=21>
                        <IMG height=19 
            src="Images/Face/face21.gif" width=19 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=22>
                        <IMG height=19 
            src="Images/Face/face22.gif" width=19 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=23>
                        <IMG height=19 
            src="Images/Face/face23.gif" width=19 align=middle border=0></td>
                    </tr>
                    <tr>
                      <td height="46" colspan="2" align="center" style="padding-left:10px"><table width="70%" height="25"  border="0" cellpadding="-2" cellspacing="-2">
						<tr>
                          <td width="21%"><img src="Images/UBB/B.gif" width="21" height="20" onClick="bold()">&nbsp;<img src="Images/UBB/I.gif" width="21" height="20" onClick="xt()">&nbsp;<img src="Images/UBB/U.gif" width="21" height="20" onClick="ul()"></td>
                          <td width="79%">����
                            <select name="font" class="wenbenkuang" id="font" onChange="sf(this.options[this.selectedIndex].value)">
                              <option value="����" selected>����</option>
                              <option value="����">����</option>
                              <option value="����">����</option>
                              <option value="����">����</option>
                            </select>
                            �ֺ�<span class="pt9">
                            <SELECT 
      name=size class="wenbenkuang" onChange="ss(this.options[this.selectedIndex].value)">
                              <OPTION value=1>1</OPTION>
                              <OPTION value=2>2</OPTION>
                              <OPTION 
        value=3 selected>3</OPTION>
                              <OPTION value=4>4</OPTION>
                              <option value="5">5</option>
                              <option value="6">6</option>
                              <option value="7">7</option>
                            </SELECT>
��ɫ <select onChange="sc(this.options[this.selectedIndex].value)" name="color" size="1" class="wenbenkuang" id="select">
            <option selected>Ĭ����ɫ</option>
            <option style="color:#FF0000" value="#FF0000">��ɫ����</option>
            <option style="color:#0000FF" value="#0000ff">��ɫ����</option>
            <option style="color:#ff00ff" value="#ff00ff">��ɫ����</option>
            <option style="color:#009900" value="#009900">��ɫ�ഺ</option>
            <option style="color:#009999" value="#009999">��ɫ��ˬ</option>
            <option style="color:#990099" value="#990099">��ɫ�н�</option>
            <option style="color:#990000" value="#990000">��ҹ�˷�</option>
            <option style="color:#000099" value="#000099">��������</option>
            <option style="color:#999900" value="#999900">�����Ʒ�</option>
            <option style="color:#ff9900" value="#ff9900">�ֽ�����</option>
            <option style="color:#0099ff" value="#0099ff">��������</option>
            <option style="color:#9900ff" value="#9900ff">��������</option>
            <option style="color:#ff0099" value="#ff0099">���İ�ʾ</option>
            <option style="color:#006600" value="#006600">ī�����</option>
            <option style="color:#999999" value="#999999">��������</option>
        </select>                           </span></td>
                        </tr>
                      </table></td>
                      </tr>
                    <tr>
                      <td height="145" align="center">���ݣ�</td>
                      <td><textarea name="content" cols="60" rows="14" class="wenbenkuang" id="content"
					   onkeydown=CountStrByte(this.form.content,this.form.total,this.form.used,this.form.remain);
					    onkeyup=CountStrByte(this.form.content,this.form.total,this.form.used,this.form.remain);></textarea></td>
                    </tr>
                    <tr>
                      <td height="33" align="center" style="padding-left:10px"><p>�ֽڣ�</p>                        </td>
                      <td style="padding-left:10px" class="word_grey">������� 
                        <input name="total" type="text" disabled class="noborder" id="total"  value="1600" size="4"> 
                        ���ֽ� �����ֽڣ�&nbsp;
                        <input name="used" type="text" disabled class="noborder" id="used"  value="0" size="4">                        
                        ʣ���ֽڣ�
                        <input name="remain" type="text" disabled class="noborder" id="remain" value="1600" size="4">
                        <input name="HIP" type="hidden" id="HIP" value="<%=Request.ServerVariables("REMOTE_ADDR")%>"></td>
                    </tr>
                    <tr>
                      <td height="34" align="center" style="padding-left:10px">&nbsp;</td>
                      <td class="word_grey"><input name="Button" type="button" class="btn_grey" value="�ύ������Ϣ" onClick="check()">                          
                        <input name="Submit2" type="reset" class="btn_grey" value="��д������Ϣ"></td></tr>
                  </table>
				    </form>
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
<%
rs.close()
Set rs=Nothing
%>
