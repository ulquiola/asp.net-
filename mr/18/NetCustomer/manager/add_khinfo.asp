<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<!--#include file="../../NetCustomer/F_Display.asp" -->
<%  '保存客户信息
if request.Form("khname")<>"" then
	khname=request.Form("khname")
	khsymbol=request.Form("khsymbol")
	khtype=request.Form("khtype")
	khgrade=request.Form("khgrade")
	sellsum=request.Form("sellsum")
	frdb=request.Form("frdb")
	qyname=request.Form("qyname")
	postcode=request.Form("postcode")
	address=request.Form("address")
	tel=request.Form("tel")
	email=request.Form("email")
	fax=request.Form("fax")
	homepage=request.Form("homepage")
	khh=request.Form("khh")
	zh=request.Form("zh")
	memo=request.Form("memo")
	Ins="Insert into DB_KhInfo (khnumber,khname,khsymbol,khtype,khgrade,sellsum,frdb,"&_
	"qyname,postcode,address,tel,email,fax,homepage,khh,zh,memo) values('"&session("autoNO")&_
	"','"&khname&"','"&khsymbol&"','"&khtype&"','"&khgrade&"',"&sellsum&",'"&frdb&"',"&_
	qyname&",'"&postcode&"','"&address&"','"&tel&"','"&email&"','"&fax&"','"&_
	homepage&"','"&khh&"','"&zh&"','"&memo&"')"
	conn.execute(Ins) %>
	  <script language="javascript">
	  alert("数据添加成功！");
	  window.close();
	  </script>
<%end if%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>客户信息！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image: url(../Images/bg.gif);
}
.style2 {color: #a2bcc5}
.style4 {color: #FFFFFF}
-->
</style></head>

<body>
<script language="javascript">
function mycheck(){
if (form1.khname.value=="")
{alert("请输入客户名称！");form1.khname.focus();return;}
if(form1.frdb.value=="")
{alert("请输入法人代表！");form1.frdb.focus();return;}
if(isNaN(form1.sellsum.value))
{alert("您输入的销售数量不正确！");form1.sellsum.focus();return;}
if(form1.tel.value=="")
{alert("请输入电话号码！");form1.tel.focus();return;}
form1.submit();
}
</script>
<%  '查询客户信息数据表中最大的客户编号
Set rs_Max = Server.CreateObject("ADODB.Recordset")
sql_max = "SELECT Max(khnumber) as MaxNO  FROM DB_khinfo"
rs_Max.open sql_max,conn,1,3
Dim no,valno,mymonth,cmonth,cday
no=trim(rs_Max("MaxNO"))
'生成五位的编号
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
intno="KH"+cno '组成新的数据编号
session("autoNO")=intno  '将生成的数据编号赋值给会话变量
%>
<table width="595" height="242" border="0" cellpadding="-2" cellspacing="-2" background="../Images/manage.gif">
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
          <form ACTION="add_khinfo.asp" METHOD="POST" name="form1">
            <table width="539" border="1" align="center" bordercolor="#669999" bordercolordark="#FFFFFF" cellpadding="-2" cellspacing="-2">
              <tr valign="baseline">
                <td width="60" align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">客户编号：</div></td>
                <td width="103">
                  <div align="left">
                    &nbsp;
                    <input name="khnumber" type="text" class="Sytle_text" readonly="yes" value="<%=session("autoNO")%>" size="32">
                  </div></td>
                <td width="70" bgcolor="#809EA4"><span class="style4">客户名称：</span></td>
                <td colspan="3"><div align="left">
                  &nbsp;
                  <input name="khname" type="text" class="Sytle_auto" value="" size="32">
                </div></td>
              </tr>
              <tr valign="baseline">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">客户简称：</div></td>
                <td valign="middle">
                  <div align="left">
                    &nbsp;
                    <input name="khsymbol" type="text" class="Sytle_text" value="" size="32">
                  </div></td>
                <td valign="middle" bgcolor="#809EA4"><span class="style4">客户类型：</span></td>
                <td width="125" valign="middle"><div align="left">
                  &nbsp;
                  <input name="khtype" type="text" class="Sytle_text" value="" size="32">
                </div></td>
                <td width="64" valign="middle" bgcolor="#809Ea4"><span class="style4">客户等级：</span></td>
                <td width="89" valign="middle"><div align="left">
                    &nbsp;
                    <select name="khgrade" id="select">
                      <option value="A" selected>A</option>
                      <option value="B">B</option>
                      <option value="C">C</option>
                      <option value="D">D</option>
                    </select>
                    (级)</div></td>
              </tr>
              <tr valign="baseline">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">销售数量：</div></td>
                <td valign="middle">
                  <div align="left">
                    &nbsp;
                    <input name="sellsum" type="text" class="Sytle_text" id="sellsum" value="0" size="32">
                  </div></td>
                <td valign="middle" bgcolor="#809EA4"><span class="style4">法人代表：</span></td>
                <td colspan="3" valign="middle"><div align="left">
                  &nbsp;
                  <input name="frdb" type="text" class="Sytle_text" value="" size="32">
                </div></td>
                </tr>
              <tr valign="baseline">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">所属区域：</div></td>
                <td colspan="5" valign="middle">
                  <div align="left">&nbsp;
				    <% qyNO=request("qyNO")
				  call Display(qyNO)
				 %>
				    <input type="hidden" name="qyname" value="<%=request("qyNO")%>">
                  </div></td>
              </tr>
              <tr valign="baseline">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">邮政编码：</div></td>
                <td valign="middle"><div align="left">
                  &nbsp;
                  <input name="postcode" type="text" class="Sytle_text" value="" size="32" maxlength="6">
                </div></td>
                <td valign="middle" bgcolor="#809EA4"><span class="style4">地　　址：</span></td>
                <td colspan="3" valign="middle"><div align="left">
                  &nbsp;
                  <input name="address" type="text" class="Sytle_auto" value="" size="32">
                </div></td>
                </tr>
              <tr valign="baseline">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">电　　话：</div></td>
                <td valign="middle">
                  <div align="left">
                    &nbsp;
                    <input name="tel" type="text" class="Sytle_text" value="" size="32">
                  </div></td>
                <td valign="middle" bgcolor="#809EA4"><span class="style4">Email：</span></td>
                <td colspan="3" valign="middle"><div align="left">
                  &nbsp;
                  <input name="email" type="text" class="Sytle_auto" value="" size="32">
                </div></td>
                </tr>
              <tr valign="baseline">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">传　　真：</div></td>
                <td valign="middle">
                  <div align="left">
                    &nbsp;
                    <input name="fax" type="text" class="Sytle_text" value="" size="32">
                  </div></td>
                <td valign="middle" bgcolor="#809EA4"><span class="style4">网　　址：</span></td>
                <td colspan="3" valign="middle"><div align="left">
                  &nbsp;
                  <input name="homepage" type="text" class="Sytle_auto" value="" size="32">
                </div></td>
                </tr>
              <tr valign="baseline">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">开 户 行：</div></td>
                <td valign="middle">
                  <div align="left">
                    &nbsp;
                    <input name="khh" type="text" class="Sytle_text" value="" size="32">
                  </div></td>
                <td valign="middle" bgcolor="#809EA4"><span class="style4">账　　号：</span></td>
                <td colspan="3" valign="middle"><div align="left">
                  &nbsp;
                  <input name="zh" type="text" class="Sytle_auto" value="" size="32">
                </div></td>
                </tr>
              <tr valign="baseline">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">备　　注：</div></td>
                <td colspan="5" valign="middle">
                  <div align="left">
                    &nbsp;
                    <input name="memo" type="text" class="Style_subject" value="" size="32">
                  </div></td>
              </tr>
              <tr valign="baseline">
                <td colspan="6" align="center" valign="middle" nowrap>
				<input type="button" class="Style_button" value="保存" onclick="mycheck()">
				<input name="Reset" type="reset" class="Style_button" value="重置">
				<input name="Submit2" type="button" class="Style_button" value="关闭"
				 onClick="JScript:window.close()"></td>
                </tr>
            </table>
          </form>
          <p>&nbsp;</p>
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

