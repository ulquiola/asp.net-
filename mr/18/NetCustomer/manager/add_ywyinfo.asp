<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<!--#include file="../../NetCustomer/F_Display.asp" -->
<%
session("userno")=session("autoNO")
if request.Form("name")<>"" then
'�ж������ҵ��Ա�Ƿ����
  Set rs_W = Server.CreateObject("ADODB.Recordset")
  sql_W="SELECT name FROM DB_WorkerInfo WHERE name='"&request.Form("name")&"'"
  rs_W.open sql_W,conn,1,3
  if rs_W.eof then  '��������������ҵ��Ա��Ϣ
  cname=Request.Form("name")
  sex=Request.Form("sex")
  folk=Request.Form("folk")
  birthday=Request.Form("birthday")
  rcdate=Request.Form("rcdate")
  xl=Request.Form("xl")
  sfzNO=Request.Form("sfzNO")
  zy=Request.Form("zy")
  tel=Request.Form("tel")
  workqy=Request.Form("workqy")
  address=Request.Form("address")
  email=Request.Form("email")
  memo=Request.Form("memo")
  Ins="Insert Into DB_WorkerInfo (ywynumber,name,sex,folk,birthday,rcdate,xl,sfzNO,"&_
  "zy,tel,workqy,address,email,memo) values('"&session("autoNO")&"','"&cname&"','"&sex&_
  "','"&folk&"','"&birthday&"','"&rcdate&"','"&xl&"','"&sfzNO&"','"&zy&"','"&tel&"','"&_
  workqy&"','"&address&"','"&email&"','"&memo&"')"
  conn.execute(Ins)
  Ins_P="Insert Into DB_Purview (UserNO) values('"&session("autoNO")&"')"
  conn.execute(Ins_P)
%>
 	<script language="javascript">
	alert("ҵ��Ա��Ϣ��ӳɹ���");
	window.close();
	</script>
<% response.end
else%>
 	<script language="javascript">
	alert("��ҵ��Ա��Ϣ�Ѿ����ڣ�");
	window.close();
	</script>
<% end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>ҵ��Ա��Ϣ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image: url(../Images/bg.gif);
}
.style1 {color: #FFFFFF}
.style2 {color: #a2bcc5}
-->
</style></head>

<body>
<script language="javascript">
function mycheck(){
if (form1.name.value=="")
{alert("������ҵ��Ա������");form1.name.focus();return;}
if(form1.sfzNO.value=="")
{alert("���������֤���룡");form1.sfzNO.focus();return;}
form1.submit();
}
</script>
<% 
'��ѯҵ��Ա��Ϣ���ݱ�������ҵ��Ա���
Set rs_Max = Server.CreateObject("ADODB.Recordset")
sql_max = "SELECT Max(ywynumber) as MaxNO  FROM DB_Workerinfo"
rs_Max.open sql_max,conn,1,3
Dim no,valno,mymonth,cmonth,cday
no=trim(rs_Max("MaxNO"))
'������λ�ı��
select case len(int(Right(no,5)+1))
	case 1
		cno="0000"+Cstr(int(Right(no,5)+1))
	case 2
		cno="000"+Cstr(int(Right(no,5)+1))
	case 3
		cno="00"+Cstr(int(Right(no,5)+1))
	case 4
		cno="0"+Cstr(int(Right(no,5)+1))
	case 5
		cno=Cstr(int(Right(no,5)+1))
	case Else
		cno="00001"
end select
intno="YWY"+cno '����µ����ݱ��
session("autoNO")=intno  '�����ɵ����ݱ�Ÿ�ֵ���Ự����
%>
<table width="595" height="242" border="0" cellpadding="-2" cellspacing="-2"
 background="../Images/manage.gif">
  <tr>
    <td valign="top"><table width="595" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="10" height="112">&nbsp;</td>
        <td width="577">&nbsp;</td>
        <td width="10"><span class="style2"></span></td>
      </tr>
      <tr>
        <td height="258">&nbsp;</td>
        <td valign="top"><div align="center">
          <form ACTION="add_ywyinfo.asp" METHOD="POST" name="form1">
            <table border="1" align="center" cellpadding="-2" cellspacing="-2"
			 bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr valign="middle">
                <td width="72" align="right" nowrap bgcolor="#809EA4">
				<div align="center" class="style1">�ࡡ���ţ�</div></td>
                <td width="105">                    <div align="left">
                  <input name="ywynumber" type="text" class="Sytle_text" id="ywynumber"
				   value="<%=session("autoNO")%>" size="32" readonly="yes">
                </div></td>
                <td width="71" bgcolor="#809EA4"><div align="center" class="style1">�ա�������</div></td>
                <td width="96"><div align="left">
                  <input name="name" type="text" class="Sytle_text" value="" size="32">
                </div></td>
                <td width="68" bgcolor="#809EA4"><div align="center" class="style1">�ԡ�����</div></td>
                <td width="97"><div align="left">
                  <select name="sex" id="select3">
                      <option value="��" selected>��</option>
                      <option value="Ů">Ů</option>
                  </select>
                </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">�񡡡��壺</div></td>
                <td><div align="left">
                  <input name="folk" type="text" class="Sytle_text" value="����" size="32">
                </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">�������գ�</div></td>
                <td><div align="left">
                  <input name="birthday" type="text" class="Sytle_text" value="<%=date()-20*365%>" size="32">
                </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">�볧���ڣ�</div></td>
                <td><div align="left">
                  <input name="rcdate" type="text" class="Sytle_text" value="<%=date()%>" size="32">
                </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">ѧ��������</div></td>
                <td><div align="left">
                  <select name="xl" id="select4">
                      <option value="��ʿ������" selected>��ʿ������</option>
                      <option value="˶ʿ">˶ʿ</option>
                      <option value="�о���">�о���</option>
                      <option value="����">����</option>
                      <option value="ר��">ר��</option>
                      <option value="���м���ר">���м���ר</option>
                      <option value="���м�����">���м�����</option>
                  </select>
                </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">���֤�ţ�</div></td>
                <td colspan="3"><div align="left">
                  <input name="sfzNO" type="text" class="Sytle_auto" size="32">
                </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">ר����ҵ��</div></td>
                <td><div align="left">
                  <select name="zy" id="select5">
                      <option value="�����" selected>�����</option>
                      <option value="ʳƷ">ʳƷ</option>
                      <option value="��е����">��е����</option>
                      <option value="ҽ��">ҽ��</option>
                      <option value="��ʿ">��ʿ</option>
                      <option value="��ʦ">��ʦ</option>
                      <option value="����">����</option>
                  </select>
                </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">��ϵ�绰��</div></td>
                <td colspan="3"><div align="left">
                  <input name="tel" type="text" class="Sytle_auto" value="" size="32">
                </div></td>
                </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">��������</div></td>
                <td colspan="5"><div align="left">                  <table width="98%"  cellspacing="-2" cellpadding="-2">
                    <tr>
                      <td width="86%">
						<% qyNO=request("qyNO")
						call Display(qyNO)  '��ʾ��������%>
					  </td>
                      <td width="14%"><input name="workqy" type="hidden" id="workqy4" value="<%=request("qyNO")%>"></td>
                    </tr>
                  </table>
				</div></td>
                </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">�ء���ַ��</div></td>
                <td colspan="5">
                  <div align="left">
                    <input name="address" type="text" class="Style_upload" value="" size="32">                  
                  </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">Email:</div></td>
                <td colspan="5">
                  <div align="left">
                    <input name="email" type="text" class="Sytle_auto" value="" size="32">                
                  </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">������ע��</div></td>
                <td colspan="5">
                  <div align="left">
                    <input name="memo" type="text" class="Style_upload" value="" size="32">                
                  </div></td>
              </tr>
              <tr valign="middle">
                <td height="35" colspan="6" nowrap><div align="center">
                    <input name="Button" type="button" class="Style_button" value="����"
					 onClick="mycheck()">                
                    <input name="Reset" type="reset" class="Style_button" value="����">
                    <input name="Submit2" type="button" class="Style_button" value="�ر�"
					 onClick="JScript:window.close()">
                </div></td>
                </tr>
            </table>
          </form>
          <p>&nbsp;</p>
        </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="21">&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>

